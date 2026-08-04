from __future__ import annotations

import sys


def main() -> int:
    if "--smoke-test" in sys.argv:
        return 0

    from excel_git_viewer.ui import create_application

    application, window = create_application()
    window.show()
    return application.exec()


if __name__ == "__main__":
    raise SystemExit(main())
