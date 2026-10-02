"""CareerTrace Pro launcher: starts the server and opens it in a native window (falls back to your browser)."""
import os, sys, time, threading, webbrowser, urllib.request
os.chdir(os.path.dirname(os.path.abspath(__file__)))
os.environ["CT_AUTORUN"] = "1"
PORT = int(os.environ.get("PORT", "8000"))


def serve():
    import uvicorn
    host = "0.0.0.0" if os.environ.get("CODESPACES") else "127.0.0.1"
    uvicorn.run("careertrace_ui:app", host=host, port=PORT, log_level="warning")


def wait():
    for _ in range(120):
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{PORT}/docs", timeout=2)
            return True
        except Exception:
            time.sleep(1)
    return False


if __name__ == "__main__":
    from dotenv import load_dotenv
    load_dotenv()
    load_dotenv(".env.pro")
    threading.Thread(target=serve, daemon=True).start()
    if not wait():
        print("Server did not start. Check the messages above.")
        sys.exit(1)
    url = f"http://127.0.0.1:{PORT}/pro"
    print("CareerTrace Pro is running at", url)
    try:
        import webview
        webview.create_window("CareerTrace Pro", url, width=1440, height=920)
        webview.start()
    except Exception as e:
        webbrowser.open(url)
        print(f"Opened in your browser (native window not available: {type(e).__name__}). Press Ctrl+C to quit.")
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            pass
