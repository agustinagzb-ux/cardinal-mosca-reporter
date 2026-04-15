from google_auth_oauthlib.flow import InstalledAppFlow
import os

os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"

flow = InstalledAppFlow.from_client_secrets_file(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "client_secrets.json"),
    scopes=["https://www.googleapis.com/auth/adwords"]
)

print("\nEsperando autorizacion en el navegador...")
print("Si no se abre solo, copia la URL que aparece y pegala en el navegador.\n")

creds = flow.run_local_server(
    port=8080,
    open_browser=True,
    success_message="Autorizacion exitosa. Podes cerrar esta pestana y volver a la terminal."
)

print("\n✓ REFRESH TOKEN:")
print(creds.refresh_token)
print("\nCopia ese token y pegalo en la conversacion con Claude.")
