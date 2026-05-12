// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0
//
// Reference exporter plugin used to smoke-test officecli's plugin discovery
// and exporter invocation paths. Implements the bare minimum of
// docs/plugin-protocol.md §5.2.
//
// Usage (matches the protocol):
//   officecli-exporter-test --info
//   officecli-exporter-test export <source> --out <target>
//
// Behaviour:
//   - `--info` prints a manifest declaring extension `.test` with `from:docx`,
//     `from:xlsx`, `from:pptx`.
//   - `export` writes a stub file at <target> containing the source path
//     and source size. That's enough for officecli to verify the file was
//     produced.

using System.Text.Json;

if (args.Length == 0 || args[0] == "--help" || args[0] == "-h")
{
    Console.Error.WriteLine("usage: officecli-exporter-test --info");
    Console.Error.WriteLine("       officecli-exporter-test export <source> --out <target>");
    return 64;
}

if (args[0] == "--info")
{
    var manifest = new
    {
        name = "officecli-exporter-test",
        version = "0.1.0",
        protocol = 1,
        kinds = new[] { "exporter" },
        extensions = new[] { ".test" },
        supports = new[] { "from:docx", "from:xlsx", "from:pptx" },
        description = "Reference dummy exporter for officecli plugin tests",
        license = "Apache-2.0",
    };
    Console.WriteLine(JsonSerializer.Serialize(manifest));
    return 0;
}

if (args[0] != "export")
{
    Console.Error.WriteLine($"unknown subcommand: {args[0]}");
    return 64;
}

if (args.Length < 4)
{
    Console.Error.WriteLine("export requires: <source> --out <target>");
    return 64;
}

var source = args[1];
string? target = null;
for (int i = 2; i < args.Length; i++)
{
    if (args[i] == "--out" && i + 1 < args.Length) { target = args[i + 1]; i++; }
}

if (target is null)
{
    Console.Error.WriteLine("missing --out");
    return 64;
}

if (!File.Exists(source))
{
    Console.Error.WriteLine($"source not found: {source}");
    return 2;
}

var info = new FileInfo(source);
File.WriteAllText(target,
    $"# officecli-exporter-test stub\n" +
    $"source: {source}\n" +
    $"source_size_bytes: {info.Length}\n" +
    $"target_format: test\n");

return 0;
