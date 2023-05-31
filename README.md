# gubbachi-automation


### What this does
Makes HTTP requests to danamojo to upload donation data fom gubbachi

### Why is it needed
With large number of donation data, it is cumbersome to go via the UI flow and fill up details for each donor. This automation fetches the donor data from an excel and uploads the data without user intervention. 

### Prerequisites
1. Windows or Mac system with python installed.
2. Create an excel file with the name `donation.xlsx` with the same column format as given in the file available with this project. Add donation info into the excel.

### How to run
1. Install Python
2. Install the dependencies from requirements.txt
    `python3 -m pip install -r requirements.txt`
    (Consider using virtual env)
5. Run the python module 
    `python3 donation-upload.py`
6. When you receive more donation data keep adding to the excel. The status column in the excel denotes whether the donation data was uploaded to danamojo or not. If the cell is having a value of 'COMPLETED' or 'FAILED', the donation data will not be updated again. 


## Bugs & Improvments
Please add feedback, feature requests and bugs to the todo file in this repo. 

