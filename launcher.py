"""Windows MCP stdio entry. Never pollute stdout."""
import os
import platform
import sys

if __name__ == '__main__':
    if os.name != 'nt' or platform.machine().lower() not in ('amd64','x86_64'):
        print('本 MCP 只支持 Windows x86_64。',file=sys.stderr)
        raise SystemExit(2)
    if sys.version_info < (3, 10):
        print('Python 3.10+ required',file=sys.stderr)
        raise SystemExit(2)
    from mcp_server import run
    run()
