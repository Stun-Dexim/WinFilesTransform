# WinFilesTransform
  How to Use
   
  
  Create or let the script create a config.ini with defaults.
  Prepare your FileTransfers.csv with at least two columns: Origin Path, Target Path, plus any other metadata columns.
  Run the script:
  
  python windows_file_transfer.py  
  Follow prompts for CSV file path and (optionally) credentials.
   
  
  Features Recap
   
  
  INI file: Auto-creation, stores defaults, can be edited.
  Windows Authentication: Prompts for credentials, uses current user if declined.
  Multi-threading: Default 4 threads (configurable).
  Timestamped Target Directory: Optional.
  Filename Sanitization: Double extension and illegal char stripping (configurable).
  Relative Path Handling: Uses base paths from INI.
  Metadata File: CSV with hyperlinks, status, and illegal char info.
