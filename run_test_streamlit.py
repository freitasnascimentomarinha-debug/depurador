import sys
import os
from streamlit.testing.v1 import AppTest

os.environ["OPENROUTER_API_KEY"] = "dummy"

# Initialize the AppTest
at = AppTest.from_file("app.py", default_timeout=30)
at.run()

pdf_path = "/workspaces/depurador/arquivos para teste/Cotação ISSARTEL - 07 de abril de 2026.pdf"

# Let's read the file as bytes
with open(pdf_path, "rb") as f:
    pdf_bytes = f.read()

# file_uploader index 0 is "Arraste os orçamentos aqui"
# Since accept_multiple_files=True, we upload a list of mock files or similar
# AppTest expects binary data uploaded, we can do it via:
# upload_file(self, file_path_or_buffer: Union[str, bytes, IO[Any]], name: Optional[str] = None, type: Optional[str] = None)
at.file_uploader[0].upload_file(pdf_bytes, name="Cotação ISSARTEL - 07 de abril de 2026.pdf").run()

print("Files uploaded status:", at.file_uploader[0].value)

# Now, trigger button 1: "🚀 Processar orçamentos"
# The buttons index: Button 1: 🚀 Processar orçamentos -> False
at.button[1].click().run()

print("Is there any exception?")
if at.exception:
    print("Exceptions:")
    for exc in at.exception:
        print(exc)
else:
    print("No exceptions detected directly in at.exception.")

# Let's print rendered markdown, info, success, warning, or error elements.
print("\n--- Rendered Markdown texts: ---")
for m in at.markdown:
    print(m.value)

print("\n--- Rendered Text elements: ---")
for t in at.text:
    print(t.value)

print("\n--- Rendered Error elements: ---")
for e in at.error:
    print(e.value)

print("\n--- Rendered Warning elements: ---")
for w in at.warning:
    print(w.value)

print("\n--- Rendered Info elements: ---")
for i in at.info:
    print(i.value)

print("\n--- Rendered Success elements: ---")
for s in at.success:
    print(s.value)

print("\n--- File outputs/etc ---")
