# 1. Connection error
There was an error connecting to Confluence. User asked to check their internet connection.

#  2. Http status [number]
Connection was OK but http status <> 200. If http status = 404, file not found, file may be missing from Confluence Word Template downloads - production page.

# 3. Download failed
Connection and http status are OK but file doesn't exist in temp directory.

# 4. Previous version removal failed
Kill old file in final dir failed. Probably attached to an open document. Told user to close all other open docs.

# 5. Previous version uninstall failed
Kill FinalPath didn't produce any errors, but the old file still remains in the final dir.

# 6. Installation failed
No errors triggered earlier but new file is not saved in final directory.

# 7. File not found
HTTP request returned 404: Page not found. Either file isn't available on Confluence or Confluence URL is wrong.

# 8: Permission denied
User needs admin permission to write to final template directory. If Mac, direct user to download from Self Service. If PC, contact IT and have them fix–users should have read/write permission to Startup and C:\ProgramDat