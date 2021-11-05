# ArduinoExcel_30
 
Arduino Excel (former Arduino Excel Commander) is a powerful interface between Arduino and MS Excel that supports data exchanging in both directions.

Excel can represent real time data from sensors or it can be used as an extern database to overcome Arduino memory limitations. The main purposes are:
•	data harvesting and consolidation
•	monitoring activities with email alerts
•	support for experiments or advanced applications (eg: robotic devices driven by Arduino)
Arduino Excel is typically used in prototypes but even in some professional applications for scientific experiments or industrial data harvesting accomplished with cheap hardware.

The main features are:
•	data writing to any worksheet / cell
•	data retrieving from any worksheet / cell
•	email sending for alarms or notifications
•	CSV files writing
Up to four Arduino can be connected at the same time thru USB ports.

The logic is built in the Arduino sketch with simple instructions like:
// write the x variable value to worksheet 'Example' range 'B5' with two digits as decimals
myExcel.write("Example", "B5", x, 2);
or
// get the value from worksheet 'Test' range 'A3' and put it in y variable
ret = myExcel.get("Test", "A3", y);
Find more documentation in the sketch supplied as example.


History: the project started in 2015, at beginning 2020 about 5000 users have worked with it especially in education but even in scientific or industrial environments. Top user countries are USA, Brazil, EU.

Coming soon: a new pro version with TCP an MQTT protocols is under study as an interface to SQL database, stay in touch.

