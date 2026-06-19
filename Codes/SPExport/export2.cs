
<Project Sdk="Microsoft.NET.Sdk">

<PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net48</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
     <LangVersion>latest</LangVersion>
    <Nullable>enable</Nullable>
  </PropertyGroup>

<ItemGroup>
  <PackageReference Include="System.Data.SqlClient" Version="4.9.1" />
</ItemGroup>

</Project>


using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

// Polyfill: records require IsExternalInit on .NET Framework 4.x
namespace System.Runtime.CompilerServices { internal static class IsExternalInit { } }

// Disaster recovery extractor for SharePoint 2016 WSS_Content databases.
// Works without any SharePoint farm installation — only requires a SQL connection.
//
// FSSHTTPB storage model (SP2016 shredded storage):
//   All file content is wrapped in FSSHTTPB (MS-FSSHTTPB) binary protocol.
//   Each document version has multiple SQL rows in DocsToStreams/DocStreams:
//     StreamId  1              = storage manifest (package header, skip)
//     StreamId  max(StreamIds) = storage index    (assembly order metadata)
//     all other StreamIds      = data element chunks (concatenate in BSN order)
//
//   The storage index contains N terminal patterns, one per data element, in FILE order.
//   Each terminal's compact-uint ZZ encodes the content size of the element it references.
//   Matching SI sizes to data element sizes (dequeued in scan order) gives assembly order.
//
//   Terminal pattern: [XX][?][00][20][03][YY][00][00][ZZ...]
//   ZZ is a compact uint64 (MS-FSSHTTPB §2.2.1.1):
//     b0 & 1 == 1        → 1-byte,  contentStart = t+9
//     b0 & 3 == 2        → 2-byte,  contentStart = t+10
//     b0 & 7 == 4        → 4-byte,  contentStart = t+12  ← critical for large elements
//     else (0 or 8-byte) → skip
//
// Usage:
//   SharePointExtract.exe                     → export ALL published docs to C:\output
//   SharePointExtract.exe --all [outDir]      → export ALL published docs
//   SharePointExtract.exe {docId} [outDir]    → export single document1

//SharePointExtract.exe {AB9FD5C8-A96E-41C3-9F9A-E5F09B9F7A16} C:\output

class Program
{
    static readonly string ConnectionString =
        @"Server=AAA\SPTEST;Database=WSS_Content;Integrated Security=True;";

    static readonly string DefaultOutDir = @"C:\output";

    const int ManifestStreamId = 1;

    static int Main(string[] args)
    {
        bool batchMode = args.Length == 0 ||
                         args[0].Equals("--all", StringComparison.OrdinalIgnoreCase);

        string outputDir = batchMode
            ? (args.Length > 1 ? args[1] : DefaultOutDir)
            : (args.Length > 1 ? args[1] : DefaultOutDir);

        Directory.CreateDirectory(outputDir);

        using var conn = new SqlConnection(ConnectionString);
        conn.Open();
        Banner($"Connected  : {conn.DataSource}  /  {conn.Database}");

        if (batchMode)
        {
            return RunBatch(conn, outputDir);
        }
        else
        {
            Guid docId = Guid.Parse(args[0]);
            var meta = QueryDocMeta(conn, docId);
            if (meta is null) { Err($"DocId {docId} not found in AllDocs."); return 1; }
            return ExtractDocument(conn, docId, meta, outputDir, flat: true);
        }
    }

    // ── Batch mode: extract every published document ──────────────────────────

    static int RunBatch(SqlConnection conn, string outputDir)
    {
        var docs = QueryAllPublishedDocs(conn);
        if (docs.Count == 0) { Warn("No published documents found in AllDocs."); return 0; }

        Banner($"Batch mode : {docs.Count} documents to extract → {outputDir}");
        Console.WriteLine();

        int ok = 0, failed = 0;
        foreach (var meta in docs)
        {
            Info($"── {meta.DirName}/{meta.LeafName}");
            try
            {
                int rc = ExtractDocument(conn, meta.Id, meta, outputDir, flat: false);
                if (rc == 0) ok++; else failed++;
            }
            catch (Exception ex)
            {
                Err($"  Unexpected error: {ex.GetType().Name}: {ex.Message}");
                failed++;
            }
            Console.WriteLine();
        }

        Banner($"Batch done : {ok} extracted, {failed} failed");
        return failed > 0 ? 1 : 0;
    }

    // ── Single-document extraction ────────────────────────────────────────────

    static int ExtractDocument(SqlConnection conn, Guid docId, DocMeta meta,
                               string outputRoot, bool flat)
    {
        Info($"File       : {meta.LeafName}");
        Info($"Path       : {meta.DirName}/{meta.LeafName}");
        Info($"UIVersion  : {meta.UIVersion}  (Level {meta.Level}: {LevelName(meta.Level)})");

        // Determine output path — flat puts everything in outputRoot; otherwise mirror SP path
        string outputPath;
        if (flat)
        {
            outputPath = Path.Combine(outputRoot, SanitizeName(meta.LeafName));
        }
        else
        {
            string relDir = string.Join(Path.DirectorySeparatorChar.ToString(),
                meta.DirName.Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries)
                    .Select(SanitizeName));
            string dir = Path.Combine(outputRoot, relDir);
            Directory.CreateDirectory(LongPath(dir));
            outputPath = Path.Combine(dir, SanitizeName(meta.LeafName));
        }

        var chunks = LoadAllChunks(conn, docId);
        if (chunks.Count == 0)
        {
            Info("  No rows in DocsToStreams — trying legacy AllDocStreams.");
            return TryLegacyExtract(conn, docId, outputPath) ? 0 : 1;
        }

        Info($"  Chunks: {chunks.Count}  (BSN order, after dedup)");

        // Raw blob dump for diagnostics
        string rawPath = outputPath + ".fsshttpb";
        byte[] fullBlob;
        using (var ms = new MemoryStream())
        {
            foreach (var c in chunks)
                if (c.Content != null) ms.Write(c.Content, 0, c.Content.Length);
            fullBlob = ms.ToArray();
        }
        File.WriteAllBytes(LongPath(rawPath), fullBlob);

        byte[]? extracted = TryFSSHTTPBReassembly(chunks);

        // Validate assembled result against expected file magic.
        // Two known scenarios where the result may not start with the correct magic:
        //   (a) Size collision: two elements share the same size and the wrong one was
        //       picked (e.g. an embedded image comes first). Fall back to raw blob scan.
        //   (b) Shredded-data layout: for large ZIP entries SharePoint stores the PK
        //       local-file header and the compressed data in separate FSSHTTPB elements.
        //       The SI may order the large data blob before other entries so the assembled
        //       output starts with raw compressed bytes (not PK). In that case, the real
        //       ZIP begins at the first PK\x03\x04 inside the assembled output.
        if (extracted != null &&
            ExtMagic.TryGetValue(Path.GetExtension(meta.LeafName).ToLowerInvariant(), out var magic) &&
            magic != null)
        {
            bool startsOk = extracted.Length >= magic.Length &&
                            extracted.Take(magic.Length).SequenceEqual(magic);
            if (!startsOk)
            {
                // For ZIP-based formats: search for PK\x03\x04 inside the assembled
                // data and try to validate a complete ZIP from that offset.
                bool recovered = false;
                if (magic.Length >= 2 && magic[0] == 0x50 && magic[1] == 0x4B)
                {
                    int pkOff = IndexOf(extracted, new byte[] { 0x50, 0x4B, 0x03, 0x04 }, 0);
                    Info($"  DBG: assembled[0]={extracted[0]:X2} pkOff={pkOff}");
                    if (pkOff > 0)
                    {
                        int zipEnd2 = FindZipEnd(extracted, pkOff);
                        Info($"  DBG: FindZipEnd({pkOff})={zipEnd2}");
                        if (zipEnd2 > pkOff)
                        {
                            Info($"  FSSHTTPB: ZIP found at +{pkOff} inside assembled data → {zipEnd2 - pkOff:N0} bytes");
                            byte[] inner = new byte[zipEnd2 - pkOff];
                            Buffer.BlockCopy(extracted, pkOff, inner, 0, inner.Length);
                            extracted = inner;
                            recovered = true;
                        }
                    }
                }
                if (!recovered)
                {
                    Info("  FSSHTTPB: result has wrong signature for extension — falling back to scan.");
                    extracted = null;
                }
            }
        }

        extracted ??= TryScanAndStrip(fullBlob);

        if (extracted is null)
        {
            Warn($"  Could not extract — raw blob at: {rawPath}");
            return 1;
        }

        // Trim any bytes that follow the ZIP End-of-Central-Directory record.
        // FSSHTTPB assembly can include SI/manifest chunks after the real EOCD,
        // producing a duplicate EOCD that causes Office to reject the file.
        if (extracted.Length >= 4 && extracted[0] == 0x50 && extracted[1] == 0x4B)
        {
            int zipEnd = FindZipEnd(extracted, 0);
            if (zipEnd > 0 && zipEnd < extracted.Length)
            {
                Info($"  ZIP: trimmed {extracted.Length - zipEnd} trailing bytes after EOCD");
                byte[] trimmed = new byte[zipEnd];
                Buffer.BlockCopy(extracted, 0, trimmed, 0, zipEnd);
                extracted = trimmed;
            }
        }

        File.Delete(LongPath(rawPath)); // remove raw blob once extraction succeeds
        File.WriteAllBytes(LongPath(outputPath), extracted);
        ValidateOutput(outputPath);
        return 0;
    }

    // ── FSSHTTPB reassembly ───────────────────────────────────────────────────

    static byte[]? TryFSSHTTPBReassembly(List<ChunkRow> chunks)
    {
        // The storage index is the chunk with the HIGHEST StreamId (manifest is always StreamId 1).
        // This varies per document (e.g., StreamId 168 or 170) depending on how many data chunks exist.
        int siStreamId = chunks
            .Where(c => c.StreamId != ManifestStreamId && c.Content != null)
            .Select(c => c.StreamId)
            .DefaultIfEmpty(0)
            .Max();

        if (siStreamId == 0) return null;

        var dataChunks = chunks
            .Where(c => c.StreamId != ManifestStreamId &&
                        c.StreamId != siStreamId &&
                        c.Content != null)
            .OrderBy(c => c.Bsn)
            .ToList();

        var siChunks = chunks
            .Where(c => c.StreamId == siStreamId && c.Content != null)
            .OrderBy(c => c.Bsn)
            .ToList();

        if (!dataChunks.Any() || !siChunks.Any())
        {
            Info("  FSSHTTPB: no data chunks or storage index — skipping.");
            return null;
        }

        // Compute chunk end positions in the concatenated data blob (needed for b0=0x00 elements)
        var chunkEnds = new int[dataChunks.Count];
        {
            int off = 0;
            for (int ci = 0; ci < dataChunks.Count; ci++)
            {
                off += dataChunks[ci].Content?.Length ?? 0;
                chunkEnds[ci] = off;
            }
        }

        byte[] data = ConcatChunks(dataChunks);
        byte[] si = ConcatChunks(siChunks);
        Info($"  FSSHTTPB: data {data.Length:N0} bytes, SI {si.Length:N0} bytes (SI StreamId={siStreamId})");

        // ── Step 1: chain through data element terminals ──────────────────────
        // ZZ compact-uint starts at t+8. Two header layouts observed:
        //   Standard (small elements):
        //     b0 & 1 == 1  → 1-byte,  contentStart = t+9
        //     b0 & 3 == 2  → 2-byte,  contentStart = t+10
        //     b0 & 7 == 4  → 4-byte,  contentStart = t+12
        //   8-byte-prefix (large binary elements, e.g., embedded images):
        //     b0 & 0x0F == 8 → skip 8 bytes to t+16, then read ZZ as 3-byte legacy:
        //       contentSize = (data[t+16]|data[t+17]<<8|data[t+18]<<16) >> 3
        //       contentStart = t+19
        var elements = new List<ElemRange>();
        int searchPos = 0;
        while (searchPos < data.Length - 11)
        {
            int t = FindTerminal(data, searchPos);
            if (t < 0) break;

            int b0zz = data[t + 8];
            int contentStart, contentSize;
            if ((b0zz & 0x01) == 1)
            {
                contentSize  = b0zz >> 1;
                contentStart = t + 9;
            }
            else if ((b0zz & 0x03) == 2)
            {
                contentSize  = (b0zz | (data[t + 9] << 8)) >> 2;
                contentStart = t + 10;
            }
            else if ((b0zz & 0x07) == 4)
            {
                // FSSHTTPB compact-uint "4-byte" encoding is actually 3-byte value:
                // (b0|b1<<8|b2<<16)>>3, content starts at t+11 (b3 is first content byte).
                // The standard MS-FSSHTTPB spec says 4-byte, but SharePoint 2016 stores
                // only 3 value bytes here — confirmed empirically: b3 is always the first
                // byte of content (e.g., 0xFF for JPEG, 0x9C for deflate-compressed XML).
                if (t + 11 > data.Length) { searchPos = t + 1; continue; }
                uint raw3 = (uint)(b0zz | (data[t+9]<<8) | (data[t+10]<<16));
                contentSize  = (int)(raw3 >> 3);
                contentStart = t + 11;
            }
            else if ((b0zz & 0x0F) == 8)
            {
                // 8-byte prefix field (large binary element) — skip 8 bytes, then read ZZ at t+16
                if (t + 19 > data.Length) { searchPos = t + 1; continue; }
                contentSize  = (data[t+16] | (data[t+17] << 8) | (data[t+18] << 16)) >> 3;
                contentStart = t + 19;
            }
            else
            {
                searchPos = t + 1; // unknown encoding — skip
                continue;
            }

            if (contentSize <= 0 || contentStart + contentSize > data.Length)
            {
                searchPos = t + 1; // bad terminal — skip and keep searching
                continue;
            }

            elements.Add(new ElemRange(contentStart, contentSize));
            searchPos = contentStart + contentSize;
        }

        if (elements.Count == 0)
        {
            Info("  FSSHTTPB: no data elements found.");
            return null;
        }
        Info($"  FSSHTTPB: found {elements.Count} standard elements in blob");

        // ── Step 1b: per-chunk scan for i+6=0x06 image elements ──────────────
        // BSN=2104-2115 use a terminal variant: blob[i+6]=0x06, blob[i+8]=0x00.
        // Scanning the concatenated blob for these would hit false positives inside
        // raw image data; scanning each SQL chunk independently avoids that.
        // Each such chunk contributes at most one element (content from offset 12
        // past the terminal to the end of the chunk).
        {
            int blobOff = 0;
            foreach (var dc in dataChunks)
            {
                int clen = dc.Content?.Length ?? 0;
                if (dc.Content != null && clen >= 12)
                {
                    byte[] cb = dc.Content;
                    bool foundTerminal = false;
                    for (int ci = 0; ci <= clen - 11; ci++)
                    {
                        if (cb[ci+2]==0x00 && cb[ci+3]==0x20 && cb[ci+4]==0x03 &&
                            cb[ci+7]==0x00 &&
                            (cb[ci+6]==0x06 || cb[ci+6]==0x02))
                        {
                            int b0p = cb[ci + 8];
                            bool is02 = cb[ci+6] == 0x02;
                            int cStart, cSize;
                            if (!is02 && b0p == 0x00)
                            {
                                // t6=0x06 only: size not explicitly encoded — rest of chunk
                                cStart = ci + 12;
                                cSize  = clen - cStart;
                                if (cSize > 0) elements.Add(new ElemRange(blobOff + cStart, cSize));
                                foundTerminal = true;
                                break;
                            }
                            else if ((b0p & 1) == 1 && ci + 9 <= clen)
                            { cSize = b0p >> 1;                                             cStart = ci + 9; }
                            else if ((b0p & 3) == 2 && ci + 10 <= clen)
                            { cSize = (b0p | (cb[ci+9] << 8)) >> 2;                        cStart = ci + 10; }
                            else if ((b0p & 7) == 4 && is02 && ci + 12 <= clen)
                            {
                                // t6=0x02: 4-byte ZZ for elements > 2MB
                                uint r4 = (uint)(b0p | (cb[ci+9]<<8) | (cb[ci+10]<<16) | (cb[ci+11]<<24));
                                cSize = (int)(r4 >> 3);                                     cStart = ci + 12;
                            }
                            else if ((b0p & 7) == 4 && !is02 && ci + 11 <= clen)
                            {
                                // t6=0x06: 3-byte ZZ
                                uint r3 = (uint)(b0p | (cb[ci+9]<<8) | (cb[ci+10]<<16));
                                cSize = (int)(r3 >> 3);                                     cStart = ci + 11;
                            }
                            else { continue; } // unknown ZZ — keep scanning
                            if (cSize > 0 && cStart + cSize <= clen)
                                elements.Add(new ElemRange(blobOff + cStart, cSize));
                            foundTerminal = true;
                            break; // one terminal element per chunk
                        }
                    }
                    _ = foundTerminal; // suppress unused warning
                }
                blobOff += clen;
            }
        }
        // ── Step 1c: per-chunk scan for CD+EOCD block ────────────────────────
        // The CD+EOCD may follow an outer FSSHTTPB header (~85-100 bytes) in its
        // own SQL chunk. Scanning the concatenated blob misses it if the outer
        // header bytes land between the CD start and EOCD. Scanning each raw SQL
        // chunk independently sidesteps that because the content is contiguous
        // within a single chunk row.
        {
            var existingStarts = new HashSet<int>(elements.Select(e => e.Start));
            int cdBlobOff = 0;
            foreach (var dc in dataChunks)
            {
                byte[] cb = dc.Content ?? Array.Empty<byte>();
                int clen = cb.Length;
                if (clen >= 90)
                {
                    // Outer header is typically 85-100 bytes; scan bytes 0-200 for PK\x01\x02.
                    int scanEnd = Math.Min(200, clen - 68);
                    for (int ci = 0; ci <= scanEnd; ci++)
                    {
                        if (cb[ci] != 0x50 || cb[ci+1] != 0x4B || cb[ci+2] != 0x01 || cb[ci+3] != 0x02) continue;
                        // CD start found; look for EOCD within same chunk.
                        for (int ei = ci + 46; ei <= clen - 22; ei++)
                        {
                            if (cb[ei]==0x50 && cb[ei+1]==0x4B && cb[ei+2]==0x05 && cb[ei+3]==0x06)
                            {
                                int cmt = cb[ei+20]|(cb[ei+21]<<8);
                                if (ei + 22 + cmt <= clen)
                                {
                                    int absStart = cdBlobOff + ci;
                                    int totalSz  = ei - ci + 22 + cmt;
                                    if (!existingStarts.Contains(absStart))
                                    {
                                        elements.Add(new ElemRange(absStart, totalSz));
                                        Info($"  FSSHTTPB: per-chunk CD+EOCD: BSN={dc.Bsn} off={ci} size={totalSz}");
                                    }
                                    break;
                                }
                            }
                        }
                        break; // only one CD per chunk
                    }
                }
                cdBlobOff += clen;
            }
        }
        Info($"  FSSHTTPB: found {elements.Count} total elements (incl. i+6=0x06 + CD chunks)");

        // ── Step 2: read assembly order from storage index ────────────────────
        // Each terminal's ZZ compact-uint = content size of the element it references, in file order.
        // [i+6] accepts 0x00 (standard) and 0x06 (large-binary variant) — mirrors FindTerminal.
        var assemblyOrder = new List<int>();
        var dbgT02Samples = new List<string>(); // raw bytes of first few t6=0x02 terminals
        for (int i = 0; i <= si.Length - 11; i++)
        {
            if (si[i + 2] == 0x00 && si[i + 3] == 0x20 &&
                si[i + 4] == 0x03 &&
                (si[i + 6] == 0x00 || si[i + 6] == 0x06 || si[i + 6] == 0x02) &&
                si[i + 7] == 0x00)
            {
                int b0si = si[i + 8];
                bool isLargeVariant = si[i + 6] == 0x02;  // t6=0x02 uses 4-byte ZZ
                int s;
                if ((b0si & 0x01) == 1)
                { s = b0si >> 1;                                                          i += 8; }
                else if ((b0si & 0x03) == 2)
                { s = (b0si | (si[i+9] << 8)) >> 2;                                      i += 9; }
                else if ((b0si & 0x07) == 4 && i + 11 < si.Length && isLargeVariant)
                {
                    // t6=0x02 (large-element variant): true 4-byte ZZ — handles sizes
                    // > 2,097,151 bytes (e.g. 2,441,990 for a 2.4 MB embedded image).
                    // b3 is part of the ZZ, NOT the first content byte.
                    uint rawsi = (uint)(b0si | (si[i+9]<<8) | (si[i+10]<<16) | (si[i+11]<<24));
                    s = (int)(rawsi >> 3);
                    if (dbgT02Samples.Count < 3)
                        dbgT02Samples.Add($"i={i} bytes=[{si[i]:X2}{si[i+1]:X2} 00 20 03 {si[i+5]:X2} 02 00 {b0si:X2} {si[i+9]:X2} {si[i+10]:X2} {si[i+11]:X2}] s(4B)={s}");
                    i += 11;
                }
                else if ((b0si & 0x07) == 4 && i + 10 < si.Length)
                {
                    // t6=0x00/0x06 (standard): 3-byte ZZ (empirically confirmed).
                    uint rawsi = (uint)(b0si | (si[i+9]<<8) | (si[i+10]<<16));
                    s = (int)(rawsi >> 3);
                    i += 10;
                }
                else if ((b0si & 0x0F) == 8 && i + 18 < si.Length)
                {
                    // 8-byte prefix field — skip 8 bytes, then read 3-byte ZZ at i+16
                    s = (si[i+16] | (si[i+17] << 8) | (si[i+18] << 16)) >> 3;
                    i += 18;
                }
                else
                {
                    if (isLargeVariant && b0si == 0x00)
                    {
                        // t6=0x02 with b0=0x00: large element (e.g. image chunk) referenced
                        // by GUID — no size encoded in SI. Insert a sentinel (-1) to reserve
                        // the correct assembly-order position; we fill it from leftover
                        // sizeQueue entries after the size-matching pass.
                        assemblyOrder.Add(-1);
                        i += 8;  // advance past examined bytes (i+0..i+8)
                    }
                    continue;
                }
                if (s > 0) assemblyOrder.Add(s);
            }
        }
        if (dbgT02Samples.Count > 0)
            foreach (var dbg in dbgT02Samples)
                Info($"  DBG t6=0x02 SI: {dbg}");

        if (assemblyOrder.Count == 0)
        {
            Info("  FSSHTTPB: no ordering entries in storage index.");
            return null;
        }
        Info($"  FSSHTTPB: SI has {assemblyOrder.Count} ordering entries");
        {
            var first10 = string.Join(", ", assemblyOrder.Take(10));
            Info($"  DBG assemblyOrder[0..9]: {first10}");
            var last5 = string.Join(", ", assemblyOrder.Skip(Math.Max(0, assemblyOrder.Count-5)));
            Info($"  DBG assemblyOrder[last5]: {last5}");
        }

        // ── Step 3: match SI sizes to elements and assemble ───────────────────
        // Use per-size queues so that when multiple elements share the same size
        // (a size collision), they are consumed in data-scan order — which matches
        // file-position order since SharePoint writes chunks sequentially.

        var sizeQueues = new Dictionary<int, Queue<ElemRange>>();
        foreach (var e in elements)
        {
            if (!sizeQueues.TryGetValue(e.Size, out var q))
                sizeQueues[e.Size] = q = new Queue<ElemRange>();
            q.Enqueue(e);
        }

        // First pass: match positive sizes; leave null slots for sentinels (-1).
        var orderedSlots = new List<ElemRange?>();
        var unmatched = new List<int>();
        int sentinelCount = 0;
        foreach (int s in assemblyOrder)
        {
            if (s == -1)
            {
                orderedSlots.Add(null);  // placeholder for a large-element sentinel
                sentinelCount++;
            }
            else if (sizeQueues.TryGetValue(s, out var q) && q.Count > 0)
            {
                orderedSlots.Add(q.Dequeue());
            }
            else
            {
                unmatched.Add(s);
                // Don't add a slot for unmatched — they collapse out of ordered
            }
        }

        // Second pass: fill sentinel slots with any remaining (unmatched) elements.
        // In PPTX structure, image chunks (large) appear FIRST in the ZIP/assembly order,
        // followed by XML slide content (small). So fill large elements from the FRONT
        // of sentinel slots and small elements from the BACK.
        if (sentinelCount > 0)
        {
            var leftoverAll = sizeQueues.Values.SelectMany(q => q)
                .OrderBy(e => e.Start)
                .ToList();
            var leftoverLarge = leftoverAll.Where(e => e.Size >  65536).ToList();
            var leftoverSmall = leftoverAll.Where(e => e.Size <= 65536).ToList();
            int filledCount = 0;

            // Fill large elements from the front (images come first in PPTX ZIP order)
            int li = 0;
            for (int j = 0; j < orderedSlots.Count && li < leftoverLarge.Count; j++)
            {
                if (orderedSlots[j] == null) { orderedSlots[j] = leftoverLarge[li++]; filledCount++; }
            }
            // Fill small elements from the back
            int lj = leftoverSmall.Count - 1;
            for (int j = orderedSlots.Count - 1; j >= 0 && lj >= 0; j--)
            {
                if (orderedSlots[j] == null) { orderedSlots[j] = leftoverSmall[lj--]; filledCount++; }
            }
            if (filledCount > 0)
                Info($"  FSSHTTPB: filled {filledCount}/{sentinelCount} sentinel slots (large:{leftoverLarge.Count} small:{leftoverSmall.Count})");
        }

        var ordered = orderedSlots.Where(e => e != null).Select(e => e!).ToList();

        if (unmatched.Count > 0)
        {
            var top10 = string.Join(", ", unmatched.Where(s => s != -1).Take(10));
            Info($"  FSSHTTPB: {unmatched.Count} unmatched SI sizes (first10): {top10}");
        }
        // Diagnostics: large elements in sizeQueues after matching
        {
            var largeOrdered = ordered.Where(e => e.Size > 100000).Select(e => e.Size).Distinct().OrderBy(s=>s).Take(5).ToList();
            var largeUnmatched = unmatched.Where(s => s > 100000).Take(5).ToList();
            var largeLeftover = sizeQueues.Where(kv => kv.Key > 100000 && kv.Value.Count > 0)
                                           .Select(kv => $"{kv.Key}x{kv.Value.Count}").Take(5).ToList();
            if (largeOrdered.Count > 0 || largeUnmatched.Count > 0 || largeLeftover.Count > 0)
            {
                if (largeOrdered.Count > 0) Info($"  DBG large in ordered: {string.Join(",", largeOrdered)}");
                if (largeUnmatched.Count > 0) Info($"  DBG large unmatched SI: {string.Join(",", largeUnmatched)}");
                if (largeLeftover.Count > 0) Info($"  DBG large leftover sizeQueues: {string.Join(",", largeLeftover)}");
            }
            // Count all SI terminal types (to detect unhandled t+6 values)
            var siTermTypes = new Dictionary<int, int>();
            for (int ii = 0; ii <= si.Length - 11; ii++)
                if (si[ii+2]==0x00 && si[ii+3]==0x20 && si[ii+4]==0x03 && si[ii+7]==0x00)
                { int t6 = si[ii+6]; siTermTypes.TryGetValue(t6, out int cc); siTermTypes[t6]=cc+1; }
            if (siTermTypes.Count > 0)
                Info($"  DBG SI term types: {string.Join(", ", siTermTypes.Select(kv => $"t6=0x{kv.Key:X2}:{kv.Value}"))}");
        }

        // Size-collision recovery: when two elements share the same size and a non-CD
        // element was dequeued ahead of the CD+EOCD element, the CD stays in leftover
        // sizeQueues.  Scan those leftovers and append the first PK\x01\x02 element so
        // ZipAwareReconstruct can locate the central directory.
        {
            bool cdInOrdered = ordered.Any(e => e.Start + 3 < data.Length &&
                data[e.Start] == 0x50 && data[e.Start+1] == 0x4B &&
                data[e.Start+2] == 0x01 && data[e.Start+3] == 0x02);
            if (!cdInOrdered)
            {
                foreach (var kvp in sizeQueues)
                {
                    bool done = false;
                    foreach (var e in kvp.Value)
                    {
                        if (e.Start + 3 < data.Length &&
                            data[e.Start] == 0x50 && data[e.Start+1] == 0x4B &&
                            data[e.Start+2] == 0x01 && data[e.Start+3] == 0x02)
                        {
                            ordered.Add(e);
                            Info($"  FSSHTTPB: size-collision recovery: appended CD+EOCD size={e.Size}");
                            done = true;
                            break;
                        }
                    }
                    if (done) break;
                }
            }
        }

        if (ordered.Count == 0)
        {
            Info("  FSSHTTPB: SI sizes don't match data elements.");
            return null;
        }


        int totalSize = ordered.Sum(e => e.Size);
        byte[] result = new byte[totalSize];
        int pos = 0;
        foreach (var e in ordered)
        {
            Buffer.BlockCopy(data, e.Start, result, pos, e.Size);
            pos += e.Size;
        }

        Info($"  FSSHTTPB: assembled {ordered.Count}/{elements.Count} elements → {totalSize:N0} bytes");

        // ── Step 4: data-order fallback when SI match rate is poor ────────────
        // When fewer than half of data elements match SI entries, try a sorted
        // data-order assembly: PK-starting elements (ZIP local-file-header entries)
        // come first, then all others — each group preserving scan order.
        // For Office Open XML files the PK entries encode all XML content and must
        // precede the raw binary image data to produce a valid ZIP stream.
        if (elements.Count > 0 && ordered.Count < elements.Count / 2)
        {
            var pkElems  = elements.Where(e => e.Start + 1 < data.Length &&
                                               data[e.Start] == 0x50 && data[e.Start + 1] == 0x4B).ToList();
            var binElems = elements.Where(e => !(e.Start + 1 < data.Length &&
                                                 data[e.Start] == 0x50 && data[e.Start + 1] == 0x4B)).ToList();
            var sortedElems = pkElems.Concat(binElems).ToList();

            int totalDo = sortedElems.Sum(e => e.Size);
            byte[] rdo = new byte[totalDo];
            int pdo = 0;
            foreach (var e in sortedElems) { Buffer.BlockCopy(data, e.Start, rdo, pdo, e.Size); pdo += e.Size; }
            Info($"  FSSHTTPB: data-order trial {sortedElems.Count} elements ({pkElems.Count} PK + {binElems.Count} bin) → {totalDo:N0} bytes");
            if (totalDo >= 2 && rdo[0] == 0x50 && rdo[1] == 0x4B)
            {
                Info("  FSSHTTPB: data-order result starts with PK — using data-order assembly");
                return rdo;
            }
        }

        // ── Step 5: ZIP-aware reconstruction ─────────────────────────────────
        // Trigger when:
        //   (a) assembled doesn't start with PK — SI placed raw data blocks first; or
        //   (b) assembled starts with PK but FindZipEnd fails — CD/EOCD are out of
        //       canonical order (e.g. SI placed EOCD before the LFH entries).
        // In both cases, ZipAwareReconstruct rebuilds the file in correct ZIP order.
        bool needsZAR = result.Length >= 4 && !(result[0] == 0x50 && result[1] == 0x4B);
        if (!needsZAR && result.Length >= 4 && result[0] == 0x50 && result[1] == 0x4B)
            needsZAR = FindZipEnd(result, 0) < 0;

        if (needsZAR)
        {
            var largeBlockAt = new Dictionary<int, int>();
            int lb = 0;
            foreach (var e in ordered)
            {
                if (e.Start + 1 < data.Length && data[e.Start] == 0x50 && data[e.Start + 1] == 0x4B)
                    break;
                if (!largeBlockAt.ContainsKey(e.Size))
                    largeBlockAt[e.Size] = lb;
                lb += e.Size;
            }
            byte[]? zr = ZipAwareReconstruct(result, largeBlockAt, ordered, data);
            if (zr != null) return zr;
        }

        return result;
    }

    // Reconstruct a ZIP whose local-file headers and compressed data are stored in
    // separate FSSHTTPB elements and assembled out of order.  Non-PK elements may be
    // large leading blocks (images/XML), medium elements interleaved among PK-header
    // elements, or CD/EOCD blocks that precede the LFH section in the SI order.
    static byte[]? ZipAwareReconstruct(byte[] assembled, Dictionary<int, int> largeBlockAt,
                                        List<ElemRange> ordered, byte[] blobData)
    {
        // ── A. One pass over ordered: classify every FSSHTTPB element ─────────────
        // nonPkAt:  size → Queue<assembledOffset>  for raw data elements (no PK sig at all)
        // lhPosEx:  filename → (assembledOffset, elementSize) for LFH elements (PK\x03\x04)
        // CD/EOCD/DD elements (PK\x01\x02, PK\x05\x06, PK\x07\x08) are ignored here —
        // their bytes are already present in 'assembled' at the correct offsets and are
        // written verbatim at the end.
        var nonPkAt = new Dictionary<int, Queue<int>>();
        var lhPosEx = new Dictionary<string, (int aOff, int elemSize)>();
        int pkStart = -1;
        {
            int aOff2 = 0;
            foreach (var e in ordered)
            {
                bool hasPkSig = e.Start + 1 < blobData.Length &&
                                blobData[e.Start] == 0x50 && blobData[e.Start+1] == 0x4B;
                bool isLFH    = hasPkSig && e.Start + 3 < blobData.Length &&
                                blobData[e.Start+2] == 0x03 && blobData[e.Start+3] == 0x04;
                if (!hasPkSig)
                {
                    // Raw data (JPEG, compressed XML, …) — potential data source for split entries
                    if (!nonPkAt.ContainsKey(e.Size)) nonPkAt[e.Size] = new Queue<int>();
                    nonPkAt[e.Size].Enqueue(aOff2);
                }
                else if (isLFH)
                {
                    if (pkStart < 0) pkStart = aOff2;
                    // Scan the entire element for ALL LFH entries it contains.
                    // A single FSSHTTPB element can pack multiple adjacent LFH headers
                    // (compound element), each without inline data.
                    int elemEnd = e.Start + e.Size;
                    int inner   = e.Start;
                    while (inner + 30 <= elemEnd &&
                           blobData[inner]   == 0x50 && blobData[inner+1] == 0x4B &&
                           blobData[inner+2] == 0x03 && blobData[inner+3] == 0x04)
                    {
                        int fnLen2 = blobData[inner+26]|(blobData[inner+27]<<8);
                        int efLen2 = blobData[inner+28]|(blobData[inner+29]<<8);
                        if (fnLen2 == 0 || inner + 30 + fnLen2 > elemEnd) break;
                        string fn2 = Encoding.UTF8.GetString(blobData, inner+30, fnLen2);
                        int innerOff = inner - e.Start;
                        if (!lhPosEx.ContainsKey(fn2))
                            lhPosEx[fn2] = (aOff2 + innerOff, e.Size - innerOff);
                        int thisHdrLen = 30 + fnLen2 + efLen2;
                        int thisCs = blobData[inner+18]|(blobData[inner+19]<<8)|
                                     (blobData[inner+20]<<16)|(blobData[inner+21]<<24);
                        inner += thisHdrLen;
                        // Skip past inline data only when the next position is NOT another LFH.
                        // (If it is another LFH, this is a compound element with no inline data.)
                        if (thisCs > 0 && inner + thisCs <= elemEnd &&
                            !(inner + 3 < elemEnd &&
                              blobData[inner]   == 0x50 && blobData[inner+1] == 0x4B &&
                              blobData[inner+2] == 0x03 && blobData[inner+3] == 0x04))
                        {
                            inner += thisCs;
                        }
                    }
                }
                // else: CD (PK\x01\x02), EOCD (PK\x05\x06), DD (PK\x07\x08) — written verbatim
                aOff2 += e.Size;
            }
        }
        if (pkStart < 0) return null;
        Info($"  DBG ZAR: pkStart={pkStart} assembled.Length={assembled.Length}");
        Info($"  DBG ZAR: nonPkAt.Count={nonPkAt.Count} lhPosEx.Count={lhPosEx.Count}");

        // ── B. Find EOCD — scan the entire assembled array backwards ──────────────
        // The EOCD can appear anywhere: at the end (normal), before pkStart, or mid-array
        // when SI places CD/EOCD elements before or between LFH elements.
        // Accept any PK\x05\x06 whose comment-length is consistent, then verify the CD.
        int eocdPos = -1;
        for (int i = assembled.Length - 22; i >= 0; i--)
        {
            if (assembled[i]   != 0x50 || assembled[i+1] != 0x4B ||
                assembled[i+2] != 0x05 || assembled[i+3] != 0x06) continue;
            int comment = assembled[i+20] | (assembled[i+21]<<8);
            if (i + 22 + comment > assembled.Length) continue;
            // Require at least one entry (cnt != 0) — avoids trivial false positives.
            // Do NOT check disk numbers or max cnt: ZIP64 uses cnt=0xFFFF as sentinel.
            int cnt = assembled[i+8]|(assembled[i+9]<<8);
            if (cnt == 0) continue;
            eocdPos = i; break;
        }
        Info($"  DBG ZAR: eocdPos={eocdPos}");
        if (eocdPos < 0)
        {
            // Search forward for ANY PK\x05\x06 occurrence (with/without validation)
            for (int i = 0; i + 21 < assembled.Length; i++)
            {
                if (assembled[i]==0x50 && assembled[i+1]==0x4B && assembled[i+2]==0x05 && assembled[i+3]==0x06)
                {
                    int c2 = assembled[i+20]|(assembled[i+21]<<8);
                    int cnt2 = assembled[i+8]|(assembled[i+9]<<8);
                    Info($"  DBG ZAR: PK05 at {i} comment={c2} fit={(i+22+c2<=assembled.Length)} cnt={cnt2}");
                    break;
                }
            }
            return null;
        }

        // ── C. Locate CD ─────────────────────────────────────────────────────────
        // Compute cdFirst from EOCD's CD-size field (canonical position: before EOCD).
        // When the SI places the CD element *after* the EOCD element in assembly order,
        // the CD appears immediately after the EOCD in assembled → check that too.
        int cdSize  = assembled[eocdPos+12]|(assembled[eocdPos+13]<<8)|
                      (assembled[eocdPos+14]<<16)|(assembled[eocdPos+15]<<24);
        int cdFirst = eocdPos - cdSize;
        bool cdValid = cdSize > 0 && cdFirst >= 0 && cdFirst + 3 < assembled.Length &&
                       assembled[cdFirst]   == 0x50 && assembled[cdFirst+1] == 0x4B &&
                       assembled[cdFirst+2] == 0x01 && assembled[cdFirst+3] == 0x02;
        if (!cdValid)
        {
            // CD after EOCD (SI ordered EOCD before CD)
            int eocdComment = assembled[eocdPos+20]|(assembled[eocdPos+21]<<8);
            int afterEocd = eocdPos + 22 + eocdComment;
            if (afterEocd + 3 < assembled.Length &&
                assembled[afterEocd]   == 0x50 && assembled[afterEocd+1] == 0x4B &&
                assembled[afterEocd+2] == 0x01 && assembled[afterEocd+3] == 0x02)
            {
                cdFirst = afterEocd;
                cdValid = true;
                // Recompute cdSize as bytes from cdFirst to eocdPos or to end of CD block
                cdSize = 0; // will be determined by parsing
            }
        }
        if (!cdValid) { Info($"  DBG ZAR: CD not found near eocdPos={eocdPos}"); return null; }
        Info($"  DBG ZAR: cdFirst={cdFirst}");

        // Parse CD entries from cdFirst; bound by cdSize when known, otherwise by assembled.Length.
        var cdList = new List<(int lhOff, int cs, string fname, int cdOff)>();
        int cpEnd   = cdSize > 0 ? cdFirst + cdSize : assembled.Length;
        int cpAfter = cdFirst;
        for (int cp = cdFirst; cp < cpEnd && cp + 46 <= assembled.Length; )
        {
            if (assembled[cp]   != 0x50 || assembled[cp+1] != 0x4B ||
                assembled[cp+2] != 0x01 || assembled[cp+3] != 0x02) break;
            int fnLen    = assembled[cp+28] | (assembled[cp+29]<<8);
            int efLen    = assembled[cp+30] | (assembled[cp+31]<<8);
            int fcLen    = assembled[cp+32] | (assembled[cp+33]<<8);
            int entryLen = 46 + fnLen + efLen + fcLen;
            if (cp + entryLen > assembled.Length) break;
            int cscd  = assembled[cp+20]|(assembled[cp+21]<<8)|(assembled[cp+22]<<16)|(assembled[cp+23]<<24);
            int lhOff = assembled[cp+42]|(assembled[cp+43]<<8)|(assembled[cp+44]<<16)|(assembled[cp+45]<<24);
            string fname = Encoding.UTF8.GetString(assembled, cp+46, fnLen);
            cdList.Add((lhOff, cscd, fname, cp));
            cp += entryLen;
            cpAfter = cp;
        }
        int cdParsedSize = cpAfter - cdFirst;  // actual bytes consumed by CD entries
        Info($"  DBG ZAR: cdList.Count={cdList.Count}");
        if (cdList.Count == 0) return null;
        cdList.Sort((a, b) => a.lhOff.CompareTo(b.lhOff));

        // ── D. Write entries in original ZIP order ────────────────────────────────
        // Data source priority:
        //   1. Inline — LFH element contains both header and compressed data
        //               (elemSize >= hdrLen + cs)
        //   2. Sequential — consume assembled[seqDataOff..seqDataOff+cs] from the
        //                   non-PK data region (handles compound data blocks where
        //                   one FSSHTTPB element covers multiple ZIP entries)
        //   3. nonPkAt[cs] — exact-size match fallback for unusual orderings
        // When the LFH header is absent from assembled entirely (large PPTX images
        // stored as raw data elements with no PK\x03\x04 prefix), synthesize the
        // LFH from the CD entry and fall through to (2)/(3) for the data.
        // Build an ordered list of non-PK elements by their assembled position so the
        // sequential reader can consume them regardless of whether they come before or
        // after the LFH compound element (handles pkStart=0 and pkStart>0 uniformly).
        var nonPkList = nonPkAt
            .SelectMany(kv => kv.Value.Select(off => (off, size: kv.Key)))
            .OrderBy(x => x.off)
            .ToList();
        int npIdx = 0;           // current element in nonPkList
        int npOff = 0;           // byte offset within current nonPkList element

        using var ms = new MemoryStream(assembled.Length);
        foreach (var (_, cs, fname, cdOff) in cdList)
        {
            bool usedSynthLFH = false;
            int hdrLen = 0;
            int lh = 0;
            int elemSize = 0;

            if (!lhPosEx.TryGetValue(fname, out var lhInfo))
            {
                // Fallback: scan assembled directly for PK\x03\x04 + matching filename.
                byte[] fnBytes = Encoding.UTF8.GetBytes(fname);
                bool recovered = false;
                for (int si = 0; si + 30 + fnBytes.Length <= assembled.Length; si++)
                {
                    if (assembled[si] != 0x50 || assembled[si+1] != 0x4B ||
                        assembled[si+2] != 0x03 || assembled[si+3] != 0x04) continue;
                    int fnLenAt = assembled[si+26] | (assembled[si+27]<<8);
                    if (fnLenAt != fnBytes.Length) continue;
                    bool fnMatch = true;
                    for (int k = 0; k < fnBytes.Length; k++)
                        if (assembled[si+30+k] != fnBytes[k]) { fnMatch = false; break; }
                    if (!fnMatch) continue;
                    int efLenAt = assembled[si+28] | (assembled[si+29]<<8);
                    int fbHdrLen = 30 + fnBytes.Length + efLenAt;
                    int afterFbHdr = si + fbHdrLen;
                    // Detect inline data by examining the bytes after the LFH header.
                    // Do NOT gate on fbCs==cs: when the data descriptor flag (bit 3) is
                    // set, the LFH stores cs=0 and the actual size comes from the CD.
                    bool inlineOk = cs > 0 && afterFbHdr + cs <= assembled.Length &&
                        !(assembled[afterFbHdr]==0x50 && assembled[afterFbHdr+1]==0x4B &&
                          (assembled[afterFbHdr+2]==0x03||assembled[afterFbHdr+2]==0x01||assembled[afterFbHdr+2]==0x05));
                    lhInfo = (si, inlineOk ? fbHdrLen + cs : fbHdrLen);
                    recovered = true;
                    break;
                }
                if (!recovered)
                {
                    // LFH not in assembled at all — synthesize from the CD entry.
                    byte[] fnBytesS = Encoding.UTF8.GetBytes(fname);
                    byte[] synthLFH = new byte[30 + fnBytesS.Length];
                    synthLFH[0]=0x50; synthLFH[1]=0x4B; synthLFH[2]=0x03; synthLFH[3]=0x04;
                    Buffer.BlockCopy(assembled, cdOff+6,  synthLFH, 4,  2); // version needed
                    Buffer.BlockCopy(assembled, cdOff+8,  synthLFH, 6,  2); // gpBitFlag
                    Buffer.BlockCopy(assembled, cdOff+10, synthLFH, 8,  2); // compression method
                    Buffer.BlockCopy(assembled, cdOff+12, synthLFH, 10, 2); // last mod time
                    Buffer.BlockCopy(assembled, cdOff+14, synthLFH, 12, 2); // last mod date
                    Buffer.BlockCopy(assembled, cdOff+16, synthLFH, 14, 4); // CRC-32
                    Buffer.BlockCopy(assembled, cdOff+20, synthLFH, 18, 4); // compressed size
                    Buffer.BlockCopy(assembled, cdOff+24, synthLFH, 22, 4); // uncompressed size
                    synthLFH[26]=(byte)fnBytesS.Length; synthLFH[27]=(byte)(fnBytesS.Length>>8);
                    synthLFH[28]=0; synthLFH[29]=0; // no extra field
                    Buffer.BlockCopy(fnBytesS, 0, synthLFH, 30, fnBytesS.Length);
                    ms.Write(synthLFH, 0, synthLFH.Length);
                    hdrLen = synthLFH.Length;
                    usedSynthLFH = true;
                }
            }

            if (!usedSynthLFH)
            {
                lh       = lhInfo.aOff;
                elemSize = lhInfo.elemSize;
                int fnLen2 = assembled[lh+26] | (assembled[lh+27]<<8);
                int efLen2 = assembled[lh+28] | (assembled[lh+29]<<8);
                hdrLen   = 30 + fnLen2 + efLen2;
                if (lh + hdrLen > assembled.Length)
                {
                    Info($"  DBG ZAR: hdrLen OOB lh={lh} hdrLen={hdrLen} fname='{fname}'");
                    return null;
                }
                ms.Write(assembled, lh, hdrLen);

                // Inline data: this LFH element already contains the compressed content.
                if (cs > 0 && elemSize >= hdrLen + cs)
                {
                    ms.Write(assembled, lh + hdrLen, cs);
                    int ddLen = elemSize - hdrLen - cs;
                    if (ddLen >= 12 && ddLen <= 20)
                        ms.Write(assembled, lh + hdrLen + cs, ddLen);
                    continue; // data consumed inline; do not touch npIdx/npOff
                }
            }

            if (cs > 0)
            {
                // Sequential read from non-PK elements in assembled-position order.
                // Works regardless of whether non-PK elements precede or follow the
                // LFH compound block (covers pkStart=0 and pkStart>0 uniformly).
                // Also handles compound data blocks where one FSSHTTPB element spans
                // multiple ZIP entries' compressed data.
                int remaining = cs;
                while (remaining > 0)
                {
                    if (npIdx >= nonPkList.Count)
                    {
                        Info($"  DBG ZAR: no data cs={cs} fname='{fname}' remaining={remaining} npIdx={npIdx}/{nonPkList.Count}");
                        return null;
                    }
                    var (npElemOff, npElemSize) = nonPkList[npIdx];
                    int avail = npElemSize - npOff;
                    int toRead = Math.Min(remaining, avail);
                    ms.Write(assembled, npElemOff + npOff, toRead);
                    remaining -= toRead;
                    npOff += toRead;
                    if (npOff >= npElemSize) { npIdx++; npOff = 0; }
                }
            }
        }

        // CD verbatim. EOCD: patch cdOffset to reflect actual CD position in the output
        // stream so that ZipFile.OpenRead can validate the central directory location.
        int eocdLen = 22 + (assembled[eocdPos+20]|(assembled[eocdPos+21]<<8));
        long newCdOff = ms.Length;
        ms.Write(assembled, cdFirst, cdParsedSize);
        byte[] eocdBytes = new byte[eocdLen];
        Buffer.BlockCopy(assembled, eocdPos, eocdBytes, 0, eocdLen);
        uint newCdOffU = (uint)newCdOff;
        eocdBytes[16] = (byte)(newCdOffU);
        eocdBytes[17] = (byte)(newCdOffU >> 8);
        eocdBytes[18] = (byte)(newCdOffU >> 16);
        eocdBytes[19] = (byte)(newCdOffU >> 24);
        ms.Write(eocdBytes, 0, eocdLen);

        byte[] reconstructed = ms.ToArray();
        Info($"  FSSHTTPB: ZIP-aware reconstruction: {cdList.Count} entries → {reconstructed.Length:N0} bytes");
        return reconstructed;
    }

    static int FindTerminal(byte[] blob, int from)
    {
        // [i+1] is intentionally not checked: DOCX/PDF have 0x00 there but PPTX has 0x03.
        for (int i = from; i <= blob.Length - 11; i++)
        {
            if (blob[i + 2] == 0x00 && blob[i + 3] == 0x20 &&
                blob[i + 4] == 0x03 &&
                blob[i + 6] == 0x00 && blob[i + 7] == 0x00)
                return i;
        }
        return -1;
    }


    static byte[] ConcatChunks(List<ChunkRow> rows)
    {
        using var ms = new MemoryStream();
        foreach (var r in rows)
            if (r.Content != null) ms.Write(r.Content, 0, r.Content.Length);
        return ms.ToArray();
    }

    // ── Signature-scan fallback ───────────────────────────────────────────────

    static readonly Dictionary<string, byte[]> ExtMagic = new()
    {
        { ".docx", new byte[] { 0x50, 0x4B } }, { ".xlsx", new byte[] { 0x50, 0x4B } },
        { ".pptx", new byte[] { 0x50, 0x4B } }, { ".xlsm", new byte[] { 0x50, 0x4B } },
        { ".zip",  new byte[] { 0x50, 0x4B } },
        { ".pdf",  new byte[] { 0x25, 0x50, 0x44, 0x46 } },
        { ".doc",  new byte[] { 0xD0, 0xCF } }, { ".xls", new byte[] { 0xD0, 0xCF } },
        { ".ppt",  new byte[] { 0xD0, 0xCF } },
        { ".png",  new byte[] { 0x89, 0x50, 0x4E, 0x47 } },
        { ".jpg",  new byte[] { 0xFF, 0xD8, 0xFF } }, { ".jpeg", new byte[] { 0xFF, 0xD8, 0xFF } },
        { ".gif",  new byte[] { 0x47, 0x49, 0x46 } },
        { ".bmp",  new byte[] { 0x42, 0x4D } },
    };

    static readonly (byte[] Magic, string Label)[] Signatures =
    {
        (new byte[] { 0x50, 0x4B, 0x03, 0x04 }, "ZIP / OpenXML (DOCX, XLSX, PPTX, ...)"),
        (new byte[] { 0x25, 0x50, 0x44, 0x46 }, "PDF"),
        (new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, "Compound File Binary (DOC, XLS, PPT)"),
        (new byte[] { 0x89, 0x50, 0x4E, 0x47 }, "PNG"),
        (new byte[] { 0xFF, 0xD8, 0xFF        }, "JPEG"),
        (new byte[] { 0x47, 0x49, 0x46        }, "GIF"),
        (new byte[] { 0x42, 0x4D              }, "BMP"),
        (new byte[] { 0x49, 0x49, 0x2A, 0x00 }, "TIFF (LE)"),
        (new byte[] { 0x4D, 0x4D, 0x00, 0x2A }, "TIFF (BE)"),
        (new byte[] { 0x37, 0x7A, 0xBC, 0xAF }, "7-Zip"),
        (new byte[] { 0x1F, 0x8B              }, "GZIP"),
    };

    static byte[]? TryScanAndStrip(byte[] blob)
    {
        foreach (var (magic, label) in Signatures)
        {
            int offset = IndexOf(blob, magic, 0);
            if (offset < 0) continue;
            Info($"  Fallback: found [{label}] at offset {offset}");

            // For ZIP files, find the End of Central Directory record so we can strip
            // any trailing garbage (e.g. the SI chunk appended after the actual ZIP).
            if (magic[0] == 0x50 && magic[1] == 0x4B)
            {
                int zipEnd = FindZipEnd(blob, offset);
                if (zipEnd > 0)
                {
                    Info($"  Fallback: ZIP EOCD at {zipEnd - 22}, trimmed to {zipEnd - offset} bytes");
                    byte[] trimmed = new byte[zipEnd - offset];
                    Buffer.BlockCopy(blob, offset, trimmed, 0, trimmed.Length);
                    return trimmed;
                }
            }

            if (offset == 0) return blob;
            byte[] result = new byte[blob.Length - offset];
            Buffer.BlockCopy(blob, offset, result, 0, result.Length);
            return result;
        }
        return null;
    }

    static int FindZipEnd(byte[] blob, int zipStart)
    {
        // Search backwards for End of Central Directory (PK\x05\x06).
        // Validate that cdOffset + cdSize == eocd_position to reject false positives
        // (compressed data can contain the 4-byte EOCD signature by coincidence).
        for (int i = blob.Length - 22; i >= zipStart; i--)
        {
            if (blob[i] != 0x50 || blob[i + 1] != 0x4B ||
                blob[i + 2] != 0x05 || blob[i + 3] != 0x06) continue;

            int commentLen = blob[i + 20] | (blob[i + 21] << 8);
            int end = i + 22 + commentLen;
            if (end > blob.Length) continue;

            // Structural check: central directory must end exactly at i.
            uint cdSize = (uint)(blob[i + 12] | (blob[i + 13] << 8) | (blob[i + 14] << 16) | (blob[i + 15] << 24));
            uint cdOffset = (uint)(blob[i + 16] | (blob[i + 17] << 8) | (blob[i + 18] << 16) | (blob[i + 19] << 24));
            if (cdOffset == 0xFFFFFFFF) continue; // ZIP64 — not handled
            if (zipStart + cdOffset + cdSize == i) return end;
        }
        return -1;
    }

    // ── Database helpers ──────────────────────────────────────────────────────

    record DocMeta(Guid Id, string LeafName, string DirName, int UIVersion, int Level);
    record ChunkRow(long Bsn, int StreamId, byte[]? Content);
    record ElemRange(int Start, int Size);

    static DocMeta? QueryDocMeta(SqlConnection conn, Guid docId)
    {
        const string sql = @"
SELECT TOP 1 Id, LeafName, DirName, UIVersion, CAST(Level AS int)
FROM   dbo.AllDocs
WHERE  Id = @id AND Type = 0
ORDER  BY UIVersion DESC";
        using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = docId;
        using var r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new DocMeta(r.GetGuid(0), r.GetString(1), r.GetString(2),
                           Convert.ToInt32(r[3]), Convert.ToInt32(r[4]));
    }

    static List<DocMeta> QueryAllPublishedDocs(SqlConnection conn)
    {
        // One row per unique document (latest published version per Id).
        // Restricts to documents that have shredded content in DocsToStreams,
        // which filters out SharePoint system files (_catalogs, master pages, etc.)
        // that are stored differently and are not user documents.
        const string sql = @"
WITH Latest AS (
    SELECT Id, LeafName, DirName, UIVersion, CAST(Level AS int) AS Level,
           ROW_NUMBER() OVER (PARTITION BY Id ORDER BY UIVersion DESC) AS rn
    FROM   dbo.AllDocs
    WHERE  Type = 0 AND Level = 1
      AND  EXISTS (SELECT 1 FROM dbo.DocsToStreams dts WHERE dts.DocId = Id)
)
SELECT Id, LeafName, DirName, UIVersion, Level
FROM   Latest
WHERE  rn = 1
ORDER  BY DirName, LeafName";
        using var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 300;
        using var r = cmd.ExecuteReader();
        var list = new List<DocMeta>();
        while (r.Read())
            list.Add(new DocMeta(r.GetGuid(0), r.GetString(1), r.GetString(2),
                                 Convert.ToInt32(r[3]), Convert.ToInt32(r[4])));
        return list;
    }

    static List<ChunkRow> LoadAllChunks(SqlConnection conn, Guid docId)
    {
        // HistVersion=0 means current published version. Filtering to it avoids mixing
        // chunks from old edit histories with the current version's chunks, which would
        // cause the wrong storage index to be selected (wrong max StreamId) and produce
        // corrupt assembly output.
        const string sql = @"
SELECT   dts.BSN, dts.StreamId, ds.Content
FROM     (SELECT DISTINCT BSN, StreamId FROM dbo.DocsToStreams WHERE DocId = @id AND HistVersion = 0) dts
JOIN     dbo.DocStreams ds ON ds.DocId = @id AND ds.BSN = dts.BSN
ORDER BY dts.BSN ASC";
        using var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 300;
        cmd.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = docId;
        using var r = cmd.ExecuteReader();
        var list = new List<ChunkRow>();
        while (r.Read())
            list.Add(new ChunkRow(Convert.ToInt64(r[0]), Convert.ToInt32(r[1]),
                                  r.IsDBNull(2) ? null : (byte[])r[2]));
        return list;
    }

    static bool TryLegacyExtract(SqlConnection conn, Guid docId, string outputPath)
    {
        const string sql = "SELECT Content FROM dbo.AllDocStreams WHERE Id = @id";
        try
        {
            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = docId;
            var content = cmd.ExecuteScalar() as byte[];
            if (content is null || content.Length == 0) { Err("  AllDocStreams: no content."); return false; }
            File.WriteAllBytes(LongPath(outputPath), content);
            ValidateOutput(outputPath);
            return true;
        }
        catch (SqlException ex) { Err($"  AllDocStreams not available: {ex.Message}"); return false; }
    }

    // ── Output validation ─────────────────────────────────────────────────────

    static void ValidateOutput(string path)
    {
        string lp = LongPath(path);
        if (!File.Exists(lp)) return;
        long size = new FileInfo(lp).Length;
        byte[] hdr = new byte[System.Math.Min(16, (int)size)];
        using (var fs = File.OpenRead(lp)) fs.Read(hdr, 0, hdr.Length);

        foreach (var (magic, label) in Signatures)
        {
            if (hdr.Length >= magic.Length && hdr.Take(magic.Length).SequenceEqual(magic))
            {
                Info($"  → {path}  ({size:N0} bytes)  [{label}]");
                return;
            }
        }
        Warn($"  → {path}  ({size:N0} bytes)  [unrecognized signature]");
    }

    // ── Utilities ─────────────────────────────────────────────────────────────

    static int IndexOf(byte[] haystack, byte[] needle, int startAt = 0)
    {
        for (int i = startAt; i <= haystack.Length - needle.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < needle.Length; j++)
                if (haystack[i + j] != needle[j]) { match = false; break; }
            if (match) return i;
        }
        return -1;
    }

    static string LevelName(int level) =>
        level switch { 1 => "Published", 2 => "Draft", _ => $"Unknown({level})" };

    // Prefix absolute paths with \\?\ to bypass the 260-char MAX_PATH limit on .NET Framework.
    // Only valid for fully-qualified Windows paths; do not use with relative paths.
    static string LongPath(string path) =>
        path.StartsWith(@"\\?\") ? path : @"\\?\" + path;

    static string SanitizeName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new StringBuilder(name.Length);
        foreach (char c in name)
            sb.Append(invalid.Contains(c) ? '_' : c);
        return sb.ToString();
    }

    static void Banner(string msg) { Console.ForegroundColor = ConsoleColor.Cyan; Console.WriteLine(msg); Console.ResetColor(); }
    static void Info(string msg) { Console.ForegroundColor = ConsoleColor.White; Console.WriteLine(msg); Console.ResetColor(); }
    static void Warn(string msg) { Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine($"[WARN]  {msg}"); Console.ResetColor(); }
    static void Err(string msg) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine($"[ERROR] {msg}"); Console.ResetColor(); }
}
