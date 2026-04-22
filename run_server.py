import uvicorn
import sys
import os
import main

if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        app_data = os.environ.get("APPDATA") or os.path.expanduser("~")
        log_dir = os.path.join(app_data, "trainform_logs")
        os.makedirs(log_dir, exist_ok=True)

        sys.stdout = open(os.path.join(log_dir, "backend_out.log"), "a", encoding="utf-8")
        sys.stderr = open(os.path.join(log_dir, "backend_err.log"), "a", encoding="utf-8")

    try:
        uvicorn.run(
            main.app,
            host="127.0.0.1",
            port=8000,
            log_level="info"
        )
    except Exception as e:
        print(f"Error fatal al iniciar servidor: {e}")
