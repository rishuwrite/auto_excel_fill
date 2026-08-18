# FBT Manifest Generator — First Use

## What you need

- A GitHub account/repository where you can upload this project.
- A Windows PC with Microsoft Excel installed.
- Your normal FedEx `booking.xls` export.

## One-time setup by the technical owner

1. Upload/push this entire project folder to your GitHub repository.
2. On GitHub, open the repository and click **Actions**.
3. Select **Build Windows FBT Generator**.
4. Click **Run workflow** → **Run workflow**.
5. Wait for the green check mark.
6. Open the completed workflow run.
7. At the bottom, under **Artifacts**, download:
   `FBT-Manifest-Generator-Windows`
8. Extract the downloaded ZIP.
9. Copy `FBT-Manifest-Generator.exe` to the Windows workstation.

The workstation user does NOT need Python, Git, LibreOffice, or GitHub.

## First actual use on the workstation

1. Double-click `FBT-Manifest-Generator.exe`.
2. Click **Browse** next to `Booking .xls`.
3. Select the FedEx booking `.xls` file.
4. Leave **NFEI mode = yes** unless your process is changed later.
5. Choose an output folder, or leave it as the same folder as the booking file.
6. Click **GENERATE FBT FILES**.
7. Open the output folder.
8. You will get only the files for countries that actually occur in the booking:
   - `FBT_US_NFEI_YES_YYYY-MM-DD.xls`
   - `FBT_NON_US_NFEI_YES_YYYY-MM-DD.xls`

## Important behavior

- `UNIT_VALUE1` is `=AE[row]/AN[row]` for US rows.
- `UNIT_VALUE1` is `=AG[row]/AN[row]` for all Non-US rows.
- Rows after the last populated field row are physically removed.
- Columns after the last populated field column are physically removed.
- The original booking file and original template files are not changed.
- Booking data is processed locally by the Windows application and is not uploaded to GitHub.

## If Windows blocks the EXE

The application itself normally does not need administrator rights when Excel is already installed.
However, a company-managed Windows workstation may block unsigned applications.

If Windows/security software blocks it:
1. Do not try to bypass the security policy.
2. Give the EXE/package to your IT or security team.
3. Ask them to approve/sign/allow the application according to company policy.

## Updating the application later

When the technical owner changes `generate_fbt.py`, `fbt_gui.py`, or the templates:

1. Push the changes to GitHub.
2. Open **Actions**.
3. Run **Build Windows FBT Generator** again.
4. Download the new artifact.
5. Replace the old EXE on the workstation.

The end user does not need to repeat the Python setup.
