# deckmint-wasm

WebAssembly bindings for [deckmint](https://crates.io/crates/deckmint), enabling PowerPoint generation directly in the browser.

```bash
cd deckmint-wasm
wasm-pack build --target web
python3 -m http.server 8080
# Open http://localhost:8080/demo/
```

See [`demo/index.html`](demo/index.html) for a working example that generates and downloads a `.pptx` from a button click.

See the [deckmint documentation](https://github.com/rebo/deckmint) for the full API.
