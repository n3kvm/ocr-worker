from msal import PublicClientApplication

CLIENT_ID = "34d39e8c-22ed-4c92-9197-177e6fbb0e4a"
TENANT_ID = "44e38316-d904-45ca-916f-1b6dd74df415"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

SCOPES = [
    "User.Read",
    "Mail.Read",
    "Mail.ReadWrite"
]
app = PublicClientApplication(
    CLIENT_ID,
    authority=AUTHORITY
)

flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" not in flow:
    raise Exception("No fue posible iniciar el Device Flow")

print(flow["message"])

result = app.acquire_token_by_device_flow(flow)

print(result)