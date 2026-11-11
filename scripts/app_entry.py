"""应用入口。WORDSHIELD_SELF_TEST=1 时只做打包自检：导入依赖并在内存中完成一次加解密，不打开界面，不写文件。"""
import os
import sys

if os.environ.get("WORDSHIELD_SELF_TEST") == "1":
    import tkinter  # noqa: F401
    import wordshield.gui  # noqa: F401
    from wordshield.crypto import decrypt_mapping, encrypt_mapping
    sample = {"version": 1, "entries": {}}
    if decrypt_mapping(encrypt_mapping(sample, "self-test-only-password"), "self-test-only-password") != sample:
        sys.exit("WordShield self-test failed")
    print("WordShield self-test OK")
    sys.exit(0)

from wordshield.gui import main

if __name__ == "__main__":
    main()
