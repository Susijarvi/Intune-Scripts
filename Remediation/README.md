# Intune Proactive Remediations

Intune Proactive Remediations are script pairs that automatically detect and fix
configuration issues on managed devices. Each package consists of:

- **Detection script** — checks whether the device is compliant
- **Remediation script** — fixes the issue if non-compliant

## How to Deploy

1. Go to **Intune** > **Devices** > **Remediations** > **Create**
2. Fill in the package name and description from the script's `.INTUNE` section
3. Upload the detection and remediation scripts
4. Configure the settings as documented in each script's `.INTUNE` block
5. Assign to device groups and set a schedule

## Available Remediations

| Folder | Description |
|--------|-------------|
| [AdminPSModules](./AdminPSModules/) | Installs and updates M365 admin PowerShell modules |
