"""
Owner-side utility. Prompts you for a new password, prints its SHA-256.

Usage:
    python make_password_hash.py
        Interactive — type the password twice (no echo), prints the hash.

    python make_password_hash.py --plain
        Same but echoes the password as you type. Useful when running over
        a remote shell where getpass() misbehaves.

Then copy the hash into auth.json's "sha256" field, commit, push.
"""

import getpass
import hashlib
import sys


def main() -> int:
    plain = "--plain" in sys.argv

    try:
        if plain:
            pw = input("Password: ")
            confirm = input("Confirm:  ")
        else:
            pw = getpass.getpass("Password: ")
            confirm = getpass.getpass("Confirm:  ")
    except (EOFError, KeyboardInterrupt):
        print("\nAborted.", file=sys.stderr)
        return 1

    if pw != confirm:
        print("Passwords don't match.", file=sys.stderr)
        return 1
    if len(pw) < 8:
        print("Password should be at least 8 chars (still hashing it for you):",
              file=sys.stderr)

    h = hashlib.sha256(pw.encode("utf-8")).hexdigest()
    print()
    print("Paste this into auth.json's sha256 field:")
    print()
    print(f"  {h}")
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
