# DummyExporter

Reference exporter plugin used to smoke-test officecli's plugin discovery and
exporter invocation paths. Not part of the main solution; built independently
when needed.

## Build

```
dotnet publish examples/plugins/dummy-exporter -c Release -o examples/plugins/dummy-exporter/out
```

## Manual smoke test

1. Build officecli main:
   `dotnet build src/officecli/officecli.csproj`
2. Build this fixture (see above).
3. Drop the binary at a plugin discovery path:
   - Windows: `mkdir %USERPROFILE%\.officecli\plugins\exporter\test`
   - copy `examples/plugins/dummy-exporter/out/officecli-exporter-test.exe`
     into that directory and rename it `plugin.exe`.
   - Linux/macOS: same idea under `~/.officecli/plugins/exporter/test/plugin`,
     `chmod +x`.
4. Verify discovery:
   `officecli plugins list` — expect `officecli-exporter-test  0.1.0  exporter  .test  <path>`
5. Verify export:
   ```
   officecli create /tmp/probe.docx --type docx
   officecli export /tmp/probe.docx --to test --out /tmp/probe.test
   cat /tmp/probe.test
   ```
   Expect a file with `# officecli-exporter-test stub` and the source path.

## Protocol coverage

This fixture exercises:
- `--info` manifest emission (§4 of the plugin protocol)
- Subprocess-based exporter invocation (§5.2)
- `(source, target)` resolution via `supports` matching (§4.4 example)
- Exit code 0 on success, 2 on corrupt input (§6.6)

It does NOT cover dump-reader or format-handler protocols. Separate fixtures
will follow when those code paths land in main.
