# Aanmelding Vrijwilliger — de Fietsboot

Multi-step aanmeldingsformulier voor vrijwilligers van de Fietsboot.

## Wat doet deze app?

1. Vrijwilliger vult 5-staps formulier in (functiekeuze, persoonsgegevens, vaarervaring, certificaten, overzicht)
2. Bij verzenden:
   - **Excel-rij** wordt toegevoegd aan `aanmeldingen-vrijwilligers.xlsx` op OneDrive
   - **E-mail naar secretaris** met PDF-bijlage (alle gegevens op 1 A4)
   - **Bevestigingsmail naar aanmelder** met dezelfde PDF

## Deployment

### 1. Maak een nieuw GitHub-repo aan

Leeg repo op `github.com/emarks1961/aanmelding-vrijwilliger`

### 2. Upload bestanden

Upload alle bestanden uit deze map naar de repo root.

### 3. Maak de Azure Static Web App aan (CLI)

```bash
az staticwebapp create \
  --name aanmelding-vrijwilliger \
  --resource-group deFietsboot \
  --source https://github.com/emarks1961/aanmelding-vrijwilliger \
  --location "westeurope" \
  --branch main \
  --login-with-github
```

### 4. Stel omgevingsvariabelen in

```bash
az staticwebapp appsettings set \
  --name aanmelding-vrijwilliger \
  --resource-group deFietsboot \
  --setting-names \
    GRAPH_TENANT_ID="53ed2d57-347d-4bb4-bb4f-7a0473fb51fc" \
    GRAPH_CLIENT_ID="9c30096f-5bcd-4044-855f-4c45b8ba90e7" \
    GRAPH_CLIENT_SECRET="<geheim>" \
    GRAPH_MAIL_FROM="admin@defietsboot.nl" \
    GRAPH_MAIL_TO="secretaris@defietsboot.nl" \
    GRAPH_ONEDRIVE_USER="admin@defietsboot.nl" \
    GRAPH_AANMELDINGEN_PATH="MS365/Aanmeldingen/aanmeldingen-vrijwilligers.xlsx"
```

### 5. Fix de GitHub Actions workflow

Pas het gegenereerde `.github/workflows/azure-static-web-apps-*.yml` aan:

```yaml
- uses: actions/checkout@v4
  with:
    submodules: true
- name: Build and Deploy
  uses: Azure/static-web-apps-deploy@v1
  with:
    azure_static_web_apps_api_token: ${{ secrets.AZURE_STATIC_WEB_APPS_API_TOKEN_... }}
    repo_token: ${{ secrets.GITHUB_TOKEN }}
    action: "upload"
    app_location: "/"
    api_location: "api"          # ← verplicht
    output_location: ""
    skip_app_build: true         # ← verplicht
```

### 6. Custom domain (optioneel)

DNS CNAME: `aanmelding` → `<azure-url>.azurestaticapps.net`

Dan in Azure Portal → Static Web App → Custom domains → Add → `aanmelding.defietsboot.nl`

## Omgevingsvariabelen

| Variable | Omschrijving |
|---|---|
| `GRAPH_TENANT_ID` | Azure AD tenant ID |
| `GRAPH_CLIENT_ID` | App registration client ID |
| `GRAPH_CLIENT_SECRET` | App registration client secret |
| `GRAPH_MAIL_FROM` | Verzendend postbus (M365) |
| `GRAPH_MAIL_TO` | Ontvanger secretaris |
| `GRAPH_ONEDRIVE_USER` | OneDrive eigenaar e-mail |
| `GRAPH_AANMELDINGEN_PATH` | Pad naar aanmeldingen Excel op OneDrive |

## Lokaal testen

```bash
cd api
npm install
# Maak local.settings.json aan met bovenstaande variabelen
func start
```

Open `index.html` in de browser (of gebruik Live Server).
