#!/usr/bin/env python3
"""
Localhost auth proxy — forwards 127.0.0.1:8000 to the Tailscale server.
Microsoft OAuth only allows HTTP redirect URIs for localhost, so this bridge
lets the browser complete the callback while the real server runs on Tailscale.
"""
import asyncio
import os

from dotenv import load_dotenv
load_dotenv("config/.env")

TARGET_HOST = os.getenv("TAILSCALE_IP", "localhost")
TARGET_PORT = 8000


async def _pipe(reader, writer):
    try:
        while not reader.at_eof():
            data = await reader.read(65536)
            if data:
                writer.write(data)
                await writer.drain()
    except Exception:
        pass
    finally:
        try:
            writer.close()
        except Exception:
            pass


async def handle_client(reader, writer):
    try:
        target_reader, target_writer = await asyncio.open_connection(TARGET_HOST, TARGET_PORT)
        await asyncio.gather(
            _pipe(reader, target_writer),
            _pipe(target_reader, writer),
        )
    except Exception:
        try:
            writer.close()
        except Exception:
            pass


async def main():
    server = await asyncio.start_server(handle_client, "127.0.0.1", TARGET_PORT)
    async with server:
        await server.serve_forever()


if __name__ == "__main__":
    asyncio.run(main())
