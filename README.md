# quasi-EDI
Using Python to simulate an EDI process.
This little Python Script acts as bridge to upload orders into our ERP system when orders came in as an csv file

there are a few things got accomplished
1) Read csv file and order details
2) simulate what a sales entry person would do when entering an order
3) archieving already entered csv orders, if there are any error, the csv file will be moved to an error folder to be looked at later
4) error checking - not assuming all data provided are correct
5) a detailed log to show what have been entered and errors
