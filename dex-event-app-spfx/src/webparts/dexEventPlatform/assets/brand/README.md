# DEX Brand Assets

Place official Deloitte brand assets here:

- `Deloitte_Logo.png` - Deloitte Logo
- `dex-orb.png` - DEX Event Experience Platform orb graphic

These assets are used in:
- Email templates (Base64 in DEX_EmailTemplates _Config Zeile)
- Outlook-Kalendereinträge
- Landing page

Note: Logos werden als Base64 in der DEX_EmailTemplates SharePoint-Liste
gespeichert (_Config Zeile: LogoBase64, DefaultImageBase64). Die Power Automate
Flows ersetzen die Platzhalter {{LOGO_URL}} und {{ORB_URL}} durch diese Base64-Werte.
