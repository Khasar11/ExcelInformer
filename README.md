# Excel Informer

This project is meant to retrieve data from databases and input them live into excel easily

## Getting Started

### Prerequisites

Excel that supports COM add ins

### Installing

Install the latest release using the setup file, open excel and open the excel informer tab, configure as like and tap "update once"
This will create a "data-ei" worksheet, here all your database connection strings will be located.

Example:
OPC example:
  Start row: 1
  End row: 2000
  Column: 3
  Offset Value: 1
  Offset Time: 6 (to leave space)
  
![image](https://github.com/Khasar11/ExcelInformer/assets/67635910/760f3384-0245-4e8f-a32f-d7361932d76f)

After scanning the scan end timestamp is inserted at the first row, column being offset time+1 
![Animation](https://github.com/Khasar11/ExcelInformer/assets/67635910/9d2a8cec-9fa5-4b32-9543-641975fe1388)

In my case im working with a WinCC OPC-UA Server for PCS 7 Siemens

## Built With

* [Convertersystems OPC-UA-CLIENT](https://github.com/convertersystems/opc-ua-client/)) - OPC UA Client ✅
* [MongoDB C# driver](https://github.com/mongodb/mongo-csharp-driver) - MongoDB Client 🚧
* Excel VSTO
