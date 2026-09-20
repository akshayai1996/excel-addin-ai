"""
Trust the self-signed cert using Windows certutil with -user flag.
Uses subprocess with input piped to auto-confirm any prompts.
"""
import subprocess, sys, os

cert_path = os.path.join(os.path.dirname(__file__), "server", "cert.pem")

# Convert PEM to DER first (certutil needs DER for -addstore without prompts on some systems)
# Use Python's ssl module to load and re-export
import ssl, tempfile, base64, re

pem_text = open(cert_path).read()
b64 = re.sub(r'-----[^-]+-----|\s', '', pem_text)
der_bytes = base64.b64decode(b64)

der_path = os.path.join(os.path.dirname(__file__), "server", "cert.der")
with open(der_path, "wb") as f:
    f.write(der_bytes)

print(f"DER cert written to: {der_path}")

# Try certutil with -user (CurrentUser store, no admin, no popup)
result = subprocess.run(
    ["certutil", "-addstore", "-user", "-f", "Root", der_path],
    capture_output=True, text=True, input="y\ny\ny\n"
)
print("STDOUT:", result.stdout)
print("STDERR:", result.stderr)
print("Return code:", result.returncode)

if result.returncode == 0:
    print("\n✅ Certificate successfully trusted in CurrentUser\\Root store!")
    print("Excel Desktop WebView2 will now accept https://localhost:3000")
else:
    print("\n❌ Failed. See output above.")
