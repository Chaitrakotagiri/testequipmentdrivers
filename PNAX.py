import pyvisa
import numpy as np
import time
from datetime import datetime
import sys
import pandas as pd
import os


class PNAX():
    def __init__(self, ip = None):

        self.rm = pyvisa.ResourceManager()
        self.channels_open = np.array([])
        self.pna = None
        self.ip = ip

    def connect(self):
        """
        Makes a connection with the given IP address using pyvisa resource manager
        If error occurs, quits the code
        """

        try:
            self.pna = self.rm.open_resource(f'TCPIP0::{self.ip}::inst0::INSTR')
            print(f"Connected on {self.ip}")
        except pyvisa.VisaIOError as e:
            print(e.args)
            print("Error connecting")
            sys.exit(0)

        return 0

    def print_id(self):
        """
        Prints the ID of the instrument that is connected
        """
        self.pna.write("*IDN?")
        print(self.pna.read())
        return 0

    def clear_measurements(self):
        '''
        Clears all the measurements present on teh window
        '''

        self.pna.write(f"CALC:PAR:DEL:ALL")
        self.channels_open = np.array([])

    def pna_reset(self):
        """
        Resets the PNA Window
        """
        self.pna.write(":SYST:FPReset")
        return 0

    def set_channel_freq(self, start_freq, stop_freq, center_freq, span, cw_freq):
        """
        Sets the frequency parameters of the channel based on the arguments below:
        start_freq: Start Frequency of the channel
        stop_freq: Stop Frequency of the channel
        center_freq: Center frequency of teh channel
        span: Span of the channel
        cw_freq: the CW signal frequency

        """
        self.pna.write(f"SENS:FREQ:STAR {start_freq}")
        self.pna.write(f"SENS:FREQ:STOP {stop_freq}")
        self.pna.write(f"SENS:FREQ:CENT {center_freq}")
        self.pna.write(f"SENS:FREQ:SPAN {span}")
        self.pna.write(f"SENS:FREQ:CW {cw_freq}")

        return 0

    def channel_power_set_ON(self):
        """
        Output power is on
        """
        self.pna.write(f"OUTP ON")

    def channel_power_set_OFF(self):
        """
        Output power is off
        """
        self.pna.write(f"OUTP OFF")

    def set_power(self, power):
        """
        Set the power level
        """
        self.pna.write(f"SOUR:POW1 {power}")

    def query_freq_start(self):
        """
        Query the system with start frequency
        """
        self.pna.write(f"SENS:OFFS:STAR?")
        print(self.pna.read())

    def query_freq_stop(self):
        """
        Query the system with stop frequency
        """
        self.pna.write(f"SENS:OFFS:STOP?")
        print(self.pna.read())

    def set_marker(self, trigger):
        if trigger:
            self.pna.write(f"DISP:WIND:ANN:MARK:STAT ON")
            self.pna.write(f"CALC1: MARK1:MAX: PEAK")
            self.pna.write(":INIT:IMM;*WAI")
            sig_dbm = self.pna.query(":CALC1:MARK1:Y?")
            sig_hz = self.pna.query(":CALC1:MARK1:X?")
        else:
            self.pna.write(f"DISP:WIND:ANN:MARK:STAT OFF")

    def marker(self, x_axis, y_axis, on=True):
        self.pna.write(f"CALC:MARK:ON")
        x_value = self.pna.write(f"CALC:MARK:X {x_axis}")
        y_value = self.pna.write(f"CALC:MARK:Y {y_axis}")
        print(x_value, y_value)
        return x_value, y_value

    def grab_screenshot(self, file_name, instr_path, pc_path):
        """
        filename : filename of the image, like test.png, test.jpeg
        instr_path: give only the folder path you want, C:\\Temp\\Screenshots  (no need file name)
        pc_path: give the pc path you want to save at, C:\\Users\\Screenshots (no need file name)
        """
        try:
            # Initiate the screenshot capture
            self.pna.write(rf":MMEM:STOR:IMAG '{instr_path}/{file_name}'")

            # Wait for the instrument to complete the screenshot capture
            time.sleep(10)

            # Check if the file exists on the instrument
            self.pna.write(rf":MMEM:CAT? '{instr_path}'")
            file_catalog = self.pna.read()

            if file_name not in file_catalog:
                raise Exception(f"The file {file_name} was not found in the directory {instr_path} on the instrument.")

            # Increase the timeout setting for reading the file
            self.pna.timeout = 60000  # Set timeout to 60 seconds

            # Read the file from the instrument in chunks to avoid timeout
            self.pna.write(rf"MMEM:TRAN? '{instr_path}/{file_name}'")

            screenshot_data = b""
            chunk_size = 1024  # 1 KB chunks

            while True:
                chunk = self.pna.read_raw(chunk_size)
                screenshot_data += chunk
                if len(chunk) < chunk_size:
                    break

            # Save the screenshot to local PC
            pc_file_name = os.path.join(pc_path, file_name)
            with open(pc_file_name, 'wb') as file:
                file.write(screenshot_data)

            print(f"Screenshot saved successfully at {pc_file_name}")

        except Exception as e:
            print(f"An error occurred: {e}")



    def get_data_output(self, param, start_freq, stop_freq, number_of_points):
        """
        param: S11, S21, S43, etc
        start_freq : start freq of the measurement
        stop_freq: stop freq of the measurement
        number of points: total number of pints between the start and stop frequency
        """

        self.pna.write("*RST")  # Reset the instrument
        self.pna.write("SYST:PRES")  # Preset the system
        self.pna.write(f"CALC:PAR:DEF 'Meas1', {param}")  # Define measurement parameter
        self.pna.write("DISP:WIND:TRAC:FEED 'Meas1'")  # Display the trace

        # Configure the frequency sweep (example)
        self.pna.write(f"SENS:FREQ:STAR {start_freq}")  # Start frequency: 1 GHz
        self.pna.write(f"SENS:FREQ:STOP {stop_freq}")  # Stop frequency: 10 GHz
        self.pna.write(f"SENS:SWE:POIN {number_of_points}")  # Number of points

        # Trigger the measurement
        self.pna.write("INIT:IMM;*WAI")

        # Retrieve the S-parameter data
        self.pna.write("FORM:DATA ASC")  # Set the data format to ASCII
        sparam_data = self.pna.query("CALC:DATA? SDATA")  # Get the S-parameter data

        # Parse the data
        data_list = [float(x) for x in sparam_data.split(',')]
        data_complex = np.array(data_list[0::2]) + 1j * np.array(data_list[1::2])

        # Save data to CSV
        df = pd.DataFrame({
            'Frequency': np.linspace(start_freq, stop_freq, number_of_points),
            'Real': data_complex.real,
            'Imaginary': data_complex.imag
        })
        df.to_csv('sparam_data.csv', index=False)

        # Save data to S2P
        with open('sparam_data.s2p', 'w') as f:
            f.write("# GHz S MA R 50\n")  # S2P file header
            for freq, real, imag in zip(np.linspace(start_freq, stop_freq, number_of_points), data_complex.real, data_complex.imag):
                f.write(f"{freq / 1e9} {real} {imag}\n")  # Frequency in GHz

        print("Done writing data to the files.")

        return 0

    def get_all_marker_data(self, file_name, trace_index):
        marker_data = []
        for i in range(1, 11):
            try:
                self.pna.write(f"CALC{trace_index}:PAR:SEL")

                self.pna.query(f"CALC{trace_index}:MARK{i}:X?")
                x_data = self.pna.read()

                self.pna.write(f"CALC{trace_index}:MARK{i}:Y?")
                y_data = self.pna.read()

                x_data_list = [float(x) for x in x_data.split(',')]
                y_data_list = [float(y) for y in y_data.split(',')]

                for x, y in zip(x_data_list, y_data_list):
                    marker_data.append({'Marker': i, 'Trace': trace_index, 'X': x, 'Y': y})
            except:
                continue

        df = pd.DataFrame(marker_data)
        df.to_excel(file_name, index=False)


    def disconnect_pna(self):
        # Close the connection
        self.pna.close()
        return 0



    """
    One Readout per Trace     DISP:WIND:ANN:MARK:SING
    Marker Readout Size       DISP:WIND:ANN:MARK:SIZE
    Measurement Trace On|Off  DISP:WIND:TRAC 
    Memory Trace On|Off       DISP:WIND:TRAC:MEM 
    Title Annotation On|Off   DISP:WIND:TITL 
    Make a Title Annotation   DISP:WIND:TITL:DATA 
    Display Update On|Off     DISP:ENAB
    Window Update On|Off      DISP:WIND:ENABle
    """


if __name__ == '__main__':
    """
    The main function used to handle/call the above functions to perform necessary testing
    """
    pnax = PNAX(ip="")  # input the IP Address
    pnax.connect()
    pnax.print_id()

