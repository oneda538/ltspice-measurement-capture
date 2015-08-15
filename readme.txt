Written by Daniel O'Neill in 2015

This application is designed to analyse LTspice IV log files and extract all 
measurements over every step. It is intended to help analyse data after 
multiple runs (i.e. Monte Carlo analysis).

By default this expects a log file in your temp directory or you can pass a 
file or folder location in as an argument. If you pass a folder it will pull 
the latest log file in that folder.

The output can either be in a csv format (You can chose a different extension) 
or in a xlsx format (default). The xlsx option has a freeze panes and auto 
filter applied. This is decided via a command line parameter.

The default run time option (for me in Win 7) is below:
LT_meas_to_csv.exe -t xlsx C:\Users\{username}\AppData\Local\Temp



How to use:

1 - Drag and drop a .log from a simulation onto the exe
2 - If your .log files are stored in your temp directory the just run the exe
3 - Run the exe in the same folder as you log file

In all cases the output will be created in the same directory as you simulation.

NOTE: You must have .step and .meas commands for this to be any use to you.
