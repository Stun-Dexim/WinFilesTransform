# WinFilesTransform
How to Use
 

	1. Install dependencies
		a. pip install pywin32  
 
	1. Place this script in a folder (e.g., windows_file_transfer.py).

	1. Prepare your FileTransfers.csv
		a. First two columns: Origin Path, Target Path
		b. Additional columns will be preserved in metadata.
		
	2. Run the script
		a. python windows_file_transfer.py  
			i. It will prompt for the CSV file path and (if needed) credentials.
			ii. If config.ini does not exist, it will offer to create one.
			
	3. Check output:
		a. Transferred files will be in the specified target locations.
Metadata CSV and a log file will be created in the same directory as your input CSV.![image](https://github.com/user-attachments/assets/761c334c-30ac-4a2d-b378-1300915714e9)
