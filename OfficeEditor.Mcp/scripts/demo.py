#!/usr/bin/env python3
"""
End-to-end demo of the OfficeEditor.Mcp stdio host.

Workflow (all over newline-delimited JSON-RPC 2.0 on stdio):
  initialize -> notifications/initialized -> tools/list
  -> deck_anatomize (upload REF deck via pptxBase64; capture deckHandle)
  -> deck_replace_element (replaceText on the first Text element of slide 1)
  -> deck_render_slide (slide 1, PNG) -> save to /tmp/officeeditor-mcp-slide1.png

Usage:  python3 scripts/demo.py
Run from anywhere; paths resolve relative to this script.
"""

import base64
import json
import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent
REPO_ROOT = PROJECT_DIR.parent
REF_DECK = REPO_ROOT / "examples" / "REF" / "PPTX" / "northwind-demo.pptx"
DLL = PROJECT_DIR / "bin" / "Debug" / "net9.0" / "OfficeEditor.Mcp.dll"
OUT_PNG = Path("/tmp/officeeditor-mcp-slide1.png")

READ_TIMEOUT_SECONDS = 600  # first render includes a cold font scan


class McpClient:
    def __init__(self, process: subprocess.Popen):
        self._process = process
        self._next_id = 1

    def request(self, method: str, params=None) -> dict:
        frame = {"jsonrpc": "2.0", "id": self._next_id, "method": method}
        self._next_id += 1
        if params is not None:
            frame["params"] = params
        self._send(frame)
        response = self._read_response()
        if "error" in response:
            raise RuntimeError(f"{method} failed: {response['error']}")
        return response["result"]

    def notify(self, method: str) -> None:
        self._send({"jsonrpc": "2.0", "method": method})

    def call_tool(self, name: str, arguments: dict) -> dict:
        result = self.request("tools/call", {"name": name, "arguments": arguments})
        if result.get("isError"):
            raise RuntimeError(f"tool {name} returned isError: {result}")
        if "structuredContent" in result:
            return result["structuredContent"]
        return json.loads(result["content"][0]["text"])

    def _send(self, frame: dict) -> None:
        line = json.dumps(frame)
        print(f">> {line[:160]}{'...' if len(line) > 160 else ''}")
        self._process.stdin.write(line + "\n")
        self._process.stdin.flush()

    def _read_response(self) -> dict:
        line = self._process.stdout.readline()
        if not line:
            stderr = self._process.stderr.read()
            raise RuntimeError(f"host closed stdout unexpectedly. stderr:\n{stderr}")
        return json.loads(line)


def main() -> int:
    if not REF_DECK.exists():
        print(f"REF deck missing: {REF_DECK}", file=sys.stderr)
        return 1

    # Build first, then run the DLL directly: `dotnet run` would print build
    # output on stdout and pollute the protocol stream.
    subprocess.run(
        ["dotnet", "build", str(PROJECT_DIR / "OfficeEditor.Mcp.csproj"), "-v", "q", "--nologo"],
        check=True, cwd=REPO_ROOT)

    deck_b64 = base64.b64encode(REF_DECK.read_bytes()).decode("ascii")
    print(f"# REF deck: {REF_DECK} ({REF_DECK.stat().st_size} bytes)")

    process = subprocess.Popen(
        ["dotnet", str(DLL)],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, bufsize=1)
    try:
        client = McpClient(process)

        init = client.request("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "demo-script", "version": "1.0"}})
        print(f"<< initialize -> server {init['serverInfo']} protocol {init['protocolVersion']}")
        client.notify("notifications/initialized")

        tools = client.request("tools/list")
        print(f"<< tools/list -> {[t['name'] for t in tools['tools']]}")

        anatomy = client.call_tool("deck_anatomize", {"pptxBase64": deck_b64})
        handle = anatomy["deckHandle"]
        print(f"<< deck_anatomize -> {anatomy['slideCount']} slides, deckHandle={handle}")
        first_slide = anatomy["slides"][0]
        for element in first_slide["elements"]:
            print(f"   slide 1 element: id={element['id']} type={element['type']} "
                  f"name={element['name']!r} text={(element.get('text') or '')[:40]!r}")

        text_element = next(e for e in first_slide["elements"] if e["type"] == "Text")
        edit = client.call_tool("deck_replace_element", {
            "deckHandle": handle,
            "operations": [{
                "type": "replaceText",
                "slide": 1,
                "elementId": text_element["id"],
                "text": "Edited via MCP"}]})
        print(f"<< deck_replace_element -> success={edit['success']} "
              f"revision={edit['revision']} changedSlides={edit['changedSlides']} "
              f"errors={edit['errors']}")
        if not edit["success"]:
            raise RuntimeError("replaceText failed")

        render = client.call_tool("deck_render_slide", {
            "deckHandle": handle, "slide": 1, "format": "png"})
        png = base64.b64decode(render["contentBase64"])
        assert png[:4] == b"\x89PNG", "rendered content is not a PNG"
        OUT_PNG.write_bytes(png)
        print(f"<< deck_render_slide -> {render['contentType']} "
              f"{len(png)} bytes, saved to {OUT_PNG}")

        print("\nDEMO OK: full MCP workflow completed.")
        return 0
    finally:
        process.stdin.close()
        process.wait(timeout=30)
        stderr = process.stderr.read()
        if stderr:
            print(f"--- host stderr ---\n{stderr}", file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main())
