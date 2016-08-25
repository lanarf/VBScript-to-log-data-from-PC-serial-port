# VBScript-to-log-data-from-serial-port

This is a little VBScript on my PC to log wind speed data from a remote wind speed sensor.  That was really the easy part; the hard part was getting the data from the sensor to the PC.  Although it has nothing at all to do with the VBScript I figured I'd detail what I had to do to get to that point.

I bought a wind speed sensor on eBay that has a 4-20ma analog output.  I had to mount the sensor on one side of my house but my PC is on the other side of the house so I needed to get the data from one side to the other.  I bought the 4-20ma version of the sensor since wire length would not be a problem as it would with the 0-5volt version (because of voltage drop).  I also figured I could actually string a twisted pair through the attic directly to the PC and bypass the network portion but that was really not practical.

In the house I converted the 4-20ma signal to 4-1/2 digit data using a SuperLogics 4014 A/D converter.  The output of the converter is RS-485 which I connected to a Lantronix SDS1100 serial-to-ethernet converter.  I connected the ethernet output from the Lantronix to a wireless bridge (DD-WRT) that is on the same LAN as my PC.

To talk to the Lantronix requires setting up a Virtual Serial Port on the PC using the Lantronix CPR Manager software.  The resulting serial port can be accessed by any PC program, just like a hardware serial port.  So to the VBScript all the hardware in the middle is transparent; it's as if the wind sensor had a serial port outputting 4-1/2 digit data.

So the only significant portion of the VBScript is using MSCOMMLib.MSComm to talk to the serial port.  The program opens the port once a minute and requests the current data from the SuperLogics by sending it the appropriate ASCII command. It then receives the data back as ASCII digits.  It repeats this 10 times every 2.5 seconds to get an average wind speed; it also takes the highest of the 10 data points as the peak wind speed. The VBScript then closes the serial port, processes the data, and waits 1 minute to repeat the process in a loop.

The 4-20ma data from the wind sensor corresponds to 0-32.4 meters/sec, which the VBScript converts to 0-72.48 mph.  It first has to subtract the 4ma value which represents 0 mph.  Everything above 4ma can be multiplied by 4.5298 to get mph.

To minimize the amount of data recorded to the log file, only wind speeds of 2 mph or peak speeds of 3 mph are counted. At midnight the max avg wind speed and peak wind speed for the day are recorded to a seperate log file and then a new daily log file is created for the next day's data.




