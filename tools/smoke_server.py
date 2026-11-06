"""Smoke test: the MCP server module loads and exposes the new operations."""
import asyncio
import sys

sys.path.insert(0, "src")


def main():
    import autocad_mcp.server as server
    from autocad_mcp.backends.com_backend import COMBackend

    ok = True

    # 1. backend surface
    for meth in ("execute_lisp", "send_command", "drawing_units", "xref_bind",
                 "layout_list", "validate_qc_report", "magicad_run"):
        has = hasattr(COMBackend, meth)
        print(f"[{'PASS' if has else 'FAIL'}] COMBackend.{meth}")
        ok &= has

    # 2. server tool registration
    for tool in ("system", "drawing", "xref", "validate", "layout", "nlp"):
        has = hasattr(server, tool)
        print(f"[{'PASS' if has else 'FAIL'}] server.{tool} registered")
        ok &= has

    # 3. dispatch: system(operation='send_command') reaches the backend
    class StubBackend:
        name = "stub"

        async def send_command(self, command):
            from autocad_mcp.backends.base import CommandResult
            return CommandResult(ok=True, payload={"echo": command})

        async def status(self):
            from autocad_mcp.backends.base import CommandResult
            return CommandResult(ok=True, payload={"backend": "stub"})

    async def run():
        server.get_backend = lambda: _awaitable(StubBackend())
        res = await server.system(operation="send_command", data={"command": "_.ZOOM\n_e\n"})
        print("[PASS] dispatch send_command ->", str(res)[:120])
        return True

    def _awaitable(value):
        fut = asyncio.get_event_loop().create_future()
        fut.set_result(value)
        return fut

    asyncio.run(run())

    print("\nRESULT:", "OK" if ok else "FAILURES")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
