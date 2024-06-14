The file paths in the python folder are sensitive and exact, changing them in any way will disrupt the entire process. That includes the completed_anomalies.txt
file; that txt file includes the powerpoint presentations that have been completed, the list is needed to stop repeated data points from occuring and to save time on execution of the program.

The program will pull tables from the power points found in the sinto analytics customer folders and put them into the anomaly
database.xlsx in _Anomaly Database folder. It excludes yearly/lifetime reports, empty tables, and powerpoints from before 2022.

Read comments inside auto_anomaly_updater.py for more information.

EDIT=========================
The file paths are sensitive but to use this script you will need to change the file paths to the correct place on your own computer. 
Additionally you will need to change the formatting of the files in order to find the correct ones for what you need. 

This was a very specific project for a specific goal. 
