import pandas as pd
import numpy as np
import time
import copy
import threading


# =============================================================================
# # GET  THE MANUFACTURING ROUTE VARIABLE
# =============================================================================
"""
This first part of the code aims to reorganize the Excel file that the company works with and adapt it for 
our application. In this Excel file, all the information for the final stage of the manufacturing of the 
threading taps appears. The purpose of this code is to obtain a variable called route (a list of lists) 
where the following information will be included: the ID (unique for each tray), the stations it needs 
to visit, and the time spent at each station. To have a clearer view of the DataFrame distribution organization 
and data, open the Excel.
"""
# Define the file path of the excel (ADD YOUR OWN FILE PATH!!!!!)
file_path = r"C:\Users\aaasc\Dropbox\TFG_LIP_Iker Ascacibar\Trabajo LIP 2025_2026\Publicación\Appendix_V1\Appendix A\Analitical model\By_invoice_order\Experimental_production_orders.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Sort by 'INVOICE ORDER' and 'Date'
df.sort_values(by=['INVOICE ORDER', 'Date'], inplace=True)

# Calculate the minimum number of trays per batch size
df['Total_Trays'] = np.ceil(df['Batch size'] / 200).astype(int)

# Create a specific ID for each tray.
for i in range(1, df['Total_Trays'].max() + 1):
    df[f'Tray_{i}'] = df['Batch size'].apply(lambda x: min(200, max(0, x - (i - 1) * 200)))
    df[f'ID_Tray_{i}'] = df.apply(lambda row: f"{row['INVOICE ORDER']}_{row['Article Code']}_{row['Batch size']}" 
                                  if row[f'Tray_{i}'] > 0 else np.nan, axis=1)

# Identify the rows in the DataFrame with the same batch size and invoice order but with a different WS to find orders that pass through more than one station, as the data is disorganized.
combined_rows = []
for (invoice_order, batch_size), group in df.groupby(['INVOICE ORDER', 'Batch size']):
    base_row = group.iloc[0].copy()
    
    # Collect the data from the additional stations the tray passes through: 'Date', 'Operation description', 'Time/tool [sec]'.
    for i in range(1, len(group)):
        next_row = group.iloc[i]
        base_row[f'Date_{i+1}'] = next_row['Date']
        base_row[f'Operation description_{i+1}'] = next_row['Operation description']
        base_row[f'Time/tool [sec]_{i+1}'] = next_row['Time/tool [sec]']
           
    combined_rows.append(base_row)

# DataFrame with combined rows
combined_df = pd.DataFrame(combined_rows)

# Rebuild the DataFrame so that each tray represents a row in the DataFrame
new_rows = []
for _, row in combined_df.iterrows():
    total_trays = int(row['Total_Trays'])
    # Reorganize the columns of the DataFrame so that each tray has all the data.
    for i in range(1, total_trays + 1):
        if not pd.isna(row[f'Tray_{i}']):
            tray_position = f"{i}/{total_trays}"
            new_row = {
                'ProductID': f"{row[f'ID_Tray_{i}']}_{tray_position}",
                'ArticleCode': row['Article Code'],
                'INVOICE ORDER': row['INVOICE ORDER'],
                'Batch size': row['Batch size'],
                'Tray_U': row[f'Tray_{i}'],
                'Tray_N': tray_position,
                'Date_1': row['Date'],
                'Operation description_1': row['Operation description'],
                'time/ws_1': row['Time/tool [sec]'] * row[f'Tray_{i}'] if pd.notna(row['Time/tool [sec]']) else np.nan,
                'Time/tool [sec]_1': row['Time/tool [sec]']
            }
            
            # If the trays go through more than one station, add new columns with the information from those stations
            for j in range(2, len(row.filter(like='Date')) + 1):
                time_tool_col = f'Time/tool [sec]_{j}'
                tray_u_value = row[f'Tray_{i}']  
                new_row[f'Date_{j}'] = row.get(f'Date_{j}', np.nan)
                new_row[f'Operation description_{j}'] = row.get(f'Operation description_{j}', np.nan)
                new_row[f'Time/tool [sec]_{j}'] = row.get(f'Time/tool [sec]_{j}', np.nan)
                
                # Check if the time column exists and calculate 'time/ws' (The time a tray spends at the station)
                if time_tool_col in row:
                    new_row[time_tool_col] = row.get(time_tool_col, np.nan)
                    new_row[f'time/ws_{j}'] = tray_u_value * row.get(time_tool_col, 0) if pd.notna(row.get(time_tool_col)) else np.nan
            
            new_rows.append(new_row)


new_df = pd.DataFrame(new_rows)

# A number is assigned to each station, and from now on, stations will be recognized by their number
operation_mapping = {
    # "Metrology (Entry) : 101"
    'Brushing': 102,
    'Brush I+D': 102,
    'Brush Deburring': 102,
    'Drag Finishing': 103,
    'Paint ': 104,
    'Laser Marking': 105
    # "Visual Inspection (Exit) : 106
}

# Create the manufacturing route list, starts with 101 and finishes in 106 always
route = []
for _, row in new_df.iterrows():
    row_route = [row['ProductID'], 101]  # Start with the product ID and the initial station 101 (Metrology Entry)

    station_times = {}  # Dictionary to accumulate total time per station

    # Loop through all operation columns 
    for i in range(1, 4):
        operation_col = f'Operation description_{i}'
        time_ws_col = f'time/ws_{i}'

        if pd.notna(row.get(operation_col)):  # If there is a valid operation
            operation_code = operation_mapping.get(row[operation_col], np.nan)  # Map the operation to its station code

            if pd.notna(operation_code):
                # Get the time spent at the station (for the number of units in the tray)
                time_ws_value = int(row[time_ws_col]) if pd.notna(row.get(time_ws_col)) else 0

                # Accumulate time per station
                if operation_code in station_times:
                    station_times[operation_code] += time_ws_value
                else:
                    station_times[operation_code] = time_ws_value

    # Append the stations and their accumulated times to the route (sorted by station number)
    for station in sorted(station_times.keys()):
        row_route.append(station)
        row_route.append(station_times[station])

    row_route.append(106)  # All trays exit through station 106 (Visual Inspection)
    route.append(row_route)

# The code sorts the route list by the invoice order, then creates a deep copy of route into routeregister
route.sort(key=lambda x: int(x[0].split('_')[0]))

routeregister = copy.deepcopy(route)

# =============================================================================
# VIRTUALIZATION OF THE STATIONS
# =============================================================================

"""
This piece of code simulates and manages the time and status of stations 102, 103, 104, and 105, 
which can be in one of three states: Busy, Finished, or Available.

 - Busy: The station is currently occupied and processing a tray. The station cannot receive a new.
 - Finished: The station has completed processing a tray, and the tray is ready to be picked up.
 - Available: The station is free and ready to accept a new tray. It has no tray currently being processed or waiting.

Each station's state and remaining processing time are stored in a shared dictionary ("self.stations"). 
When a station is occupied, it is marked as Busy and a separate thread is launched to simulate the passage of time. 
After this simulated time, the station state is automatically updated to Finished.

The "current_time" variable represents the estimated total time required to complete all the production processes 
for the desired order. It takes into account both the time during which stations are occupied 
and the time spent on robot movements between stations.
"""

class STATIONSandTIME_Manager:
    def __init__(self, acceleration= 500):
        # Initialize the simulation time
        self.current_time = 0
        
        # Lock to prevent race conditions in multithreaded access
        self.lock = threading.Lock()
        
        # Time acceleration factor to speed up or slow down simulation
        self.acceleration = acceleration
        
        # Time tracking for specific stations
        self.station102 = 0
        self.station103 = 0
        self.station104 = 0
        self.station105 = 0
        
        # Total time spent in robot movement
        self.robotmovtime = 0

        # Dictionary to store the state and occupation time of each station
        self.stations = {
            101: {"state": "Available", "Busy_time": 0},
            102: {"state": "Available", "Busy_time": 0},
            103: {"state": "Available", "Busy_time": 0},
            104: {"state": "Available", "Busy_time": 0},
            105: {"state": "Available", "Busy_time": 0},
            106: {"state": "Available", "Busy_time": 0}
        }

    def wait(self, duration):
        with self.lock:
            # Add movement duration to robot's movement time
            self.robotmovtime = self.robotmovtime + duration
            
            # Check if any station is currently busy
            busy_times = [s["Busy_time"] for s in self.stations.values() if s["state"] == "Busy"]
            any_busy = bool(busy_times)
            
            
            # If no station is busy, advance the global simulation time
            if not any_busy:
                self.current_time = self.current_time + duration
                return 

    def CheckStationState(self, station_id):
        with self.lock:
            if station_id not in self.stations:
                
                # Check if the requested station exists
                return "Station does not exist"
            
            # Return the current state of the station (Available, Busy, Finished)
            return self.stations[station_id]["state"]

 
    def OccupyStation(self, station_id, duration=0, pickup=0):
        with self.lock:
            
            # Add processing duration to the corresponding station counter
            if station_id==102:
                self.station102 = self.station102 + duration
            elif station_id==103:
                self.station103 = self.station103 + duration
            elif station_id==104:
                self.station104 = self.station104 + duration
            elif station_id==105:
                self.station105 = self.station105 + duration
                
            # Check for other busy stations to determine how to update tim
            busy_times = [s["Busy_time"] for s in self.stations.values() if s["state"] == "Busy"]
            any_busy = bool(busy_times)
            
            if not any_busy:
                # No other stations busy, just add the duration
                self.current_time += duration
            else:
                max_busy_time = max(busy_times)
                if duration > (max_busy_time * self.acceleration):
                    self.current_time += (duration - (max_busy_time * self.acceleration))
    
            # If there's a valid duration, mark the station as busy
            if duration > 0:
                self.stations[station_id]["state"] = "Busy"
                self.stations[station_id]["Busy_time"] = duration / self.acceleration
                print(f"Station {station_id} set to Busy for {duration} seconds")
                
                # Thread to automatically release the station after duration ends
                def release():
                    # Simulate the wait using real time divided by acceleration
                    time.sleep(duration / self.acceleration)
                    with self.lock:
                        # Only update if the station is still marked as Busy
                        if self.stations[station_id]["state"] == "Busy":
                            self.stations[station_id]["state"] = "Finished"
                            self.stations[station_id]["Busy_time"] = 0
                            print(f"Station {station_id} set to Finished")
                            
                # Start the release thread
                threading.Thread(target=release, daemon=True).start()
            
            # If pickup is None, release the station manually (used for manual override)
            elif pickup is None:  
                self.stations[station_id]["state"] = "Available"
                self.stations[station_id]["Busy_time"] = 0
                print(f"Station {station_id} set to Available")
        
           
manager = STATIONSandTIME_Manager(acceleration= 500)

# =============================================================================
# ROBOT MOVEMENTS TIMES
# =============================================================================   
"""
This part of the code sets the robot’s initial position at station 101 and defines the "movement_times" dictionary, 
which stores the time it takes for the robot to move between any two stations (from 101 to 106). 

These predefined movement durations are essential for calculating realistic robot travel times 
during the production process and are used to update the total processing time accordingly.
"""
robot_position = 101

movement_times = {
    (101, 101): 27, (102, 101): 46,  (103, 101): 39,  (104, 101): 34,  (105, 101): 30,  (106, 101): 31,
    (101, 102): 84, (102, 102): 84,  (103, 102): 73,  (104, 102): 90,  (105, 102): 86,  (106, 102): 82,
    (101, 103): 69, (102, 103): 75,  (103, 103): 69,  (104, 103): 86,  (105, 103): 82,  (106, 103): 76,    
    (101, 104): 65, (102, 104): 75,  (103, 104): 77,  (104, 104): 77,  (105, 104): 65,  (106, 104): 78,
    (101, 105): 72, (102, 105): 91,  (103, 105): 84,  (104, 105): 74,  (105, 105): 74,  (106, 105): 75,
    (101, 106): 36, (102, 106): 49,  (103, 106): 42,  (104, 106): 42,  (105, 106): 38,  (106, 106): 27,

}



# =============================================================================
# EXECUTE TASK_QUEUE
# =============================================================================   
"""
Processes the next task in the queue, simulates robot movement, and updates its position. 
"""
def ExecuteTaskQ():
    global robot_position
    if len(task_queue)>0:
        # task_queue_record stores a copy of tasks for testing since task_queue changes constantly and empties when all trays are processed.
        task_queue_record.append(copy.deepcopy(task_queue[0]))     
        # Take the next task from the queue
        current_task = str(task_queue.pop(0))
        move_time = movement_times.get((int(robot_position), int(current_task)), 45)
        print(f"Moving TX200 robot from station {robot_position} to {current_task} (time: {move_time}s)")
        manager.wait(float(move_time))
        robot_position = copy.deepcopy(current_task)

# =============================================================================
# CREATE TASK_QUEUE VARIABLE
# =============================================================================
"""
This piece of code manages the order in which the robot moves. It uses the "trays" variable, which is a list
containing the trays currently at various stations. "Trays" is derived from the "route" variable, which 
provides the tray ID, the stations it needs to pass through, and the time it takes at each station. These
variables are strategically retrieved using the functions OccupyStation, release_station, and CheckStationState, 
allowing the robot to modify the station states during its movement. The robot always knows the current state
of each station before making a move, enabling it to anticipate and manage its movements strategically.
"""

# Initialize the "tray"s variable
trays = [[], [], [],]

#The tray variable is cleared as the robot moves between stations, so a separate variable is needed to track each tray's location. 
#The lists in tray_register are only cleared once the last tray reaches station 106
tray_register = [[], [], []]

TrayState = "empty" 
# The task_queue variable is initialized. 
task_queue=[]

# The task_queue_register variable is created to keep track and check what happens during testing.
task_queue_record=[]

# Starts the main loop that runs while there are routes, trays, or tasks in the queue
while route or any(tray for tray in trays) or len(task_queue)>0:
    if TrayState != "empty":
       
        # This for loop manages the robot's movements by prioritizing trays that are already inside a station before introducing a new one.
        for tray in trays:
            if len(tray)>1:
                list_index = trays.index(tray)
                element_index = tray_register[list_index].index(tray[1])
                if (len(tray) > 1 
                and tray[1] in [105, 104, 103, 102] # Check if the station where the robot is going to move in the list [105, 104, 103, 102]
                and manager.CheckStationState(tray[1]) == "Available"): # Check if the station where the robot is going to move is Available
                    if (manager.CheckStationState(tray_register[list_index][element_index-2]) == "Finished" or "Available"
                    and manager.CheckStationState(tray[1]) == "Available"):
                        task_queue.append(tray_register[list_index][element_index-2]) # Add the previous tray task to the task queue
                        ExecuteTaskQ() # Execute the task from the queue
                        manager.OccupyStation(tray_register[list_index][element_index-2], 0,  pickup=None) # Turn the previous station as Available

                        # Add the current tray task to the task queue and execute it    
                        task_queue.append(tray[1])
                        ExecuteTaskQ()
                        # Occupy the station for the current tray
                        manager.OccupyStation(tray[1],tray[2])
                        # Remove the processed elements from the tray
                        tray.pop(1)
                        tray.pop(1)
                        break # Exit the loop to process the next tray
                        
                    
                if (len(tray) > 1 
                and tray[1]==106 # Check if the tray has passed all stations and only has the exit left.
                and manager.CheckStationState(tray_register[list_index][element_index-2])=="Finished"):  # Check if the tray finished the task at the current station before moving to 106
                    task_queue.append(tray_register[list_index][element_index-2]) # The station where the tray is located is added to the task_queue.
                    ExecuteTaskQ() # Execute the task from the queue
                    manager.OccupyStation(tray_register[list_index][element_index-2], 0,  pickup=None) # Turn the previous station as Available
                    task_queue.append(tray[1]) # The next station the tray needs to travel (106) to is added to the task_queue
                    ExecuteTaskQ()
                    trays.remove(tray) # The tray is removed from the list trays
                    tray_register.remove(tray_register[list_index]) # The tray is removed from the list tray_register
                    TrayState = "trying to introduce new tray" 
                    break # Exit the loop to process the next tray
                
                # If the tray first movement is at station 101, the station where is going is Available 
                if (len(tray)> 1  
                and tray[1] == 101 
                and manager.CheckStationState(tray[2]) == "Available"):
                    task_queue.append(tray[1]) # Add the current tray task to the task queue (101)
                    tray.pop(1) # Remove the first element of the tray
                    ExecuteTaskQ() 
                    task_queue.append(tray[1]) # Add The next station the tray needs to go to is added to the task_queue.
                    ExecuteTaskQ()
                    manager.OccupyStation(tray[1],tray[2]) # Occupy the station for the current tray
                    tray.pop(1)
                    tray.pop(1)
                    break # Exit the loop to process the next tray
                else:
                    TrayState = "trying to introduce new tray"   

        
    # This conditional checks if a new tray can be added when there are 4 or fewer trays, and the tray state is "trying to create new tray" or "starting."     
    if TrayState == "trying to introduce new tray" or "empty" and len(trays) <= 4 :
        for i in range(len(trays)):
            tray = trays[i]
            if not tray and route:
                if manager.CheckStationState(route[0][2]) == "Available": # Checks if the tray is empty and if the station in the route is Available
                    #Tray is updated in tray_register, and the route is assigned to the tray.
                    tray_register[i] = copy.deepcopy(route[0])
                    tray_register.append([])
                    trays[i] = route.pop(0)
                    trays.append([])
                    TrayState = "started" # The tray state is then set to "started.
                    break
            else:
                TrayState = "cannot introduce new tray" # If the conditions are not met

# =============================================================================
# 
# =============================================================================
"""
This section displays a summary of the simulation results, including the total estimated production time, 
the total working time of each individual station (102 to 105), and the robot’s total movement time. 
All times are shown in hours, and the percentage indicates how much of the total production time 
each component was active.
"""

# Prints the total estimated production time in hours
print("Current time:", round(((manager.current_time) / 60) / 60, 2), "h")

# Prints total working time of station 102 in hours and as a percentage of total time
print("102 Station total working time:", round((manager.station102 / 60) / 60, 2), "h", 
      "(", round((manager.station102 / manager.current_time) * 100, 2), "%)")

# Same for station 103
print("103 Station total working time:", round((manager.station103 / 60) / 60, 2), "h", 
      "(", round((manager.station103 / manager.current_time) * 100, 2), "%)")

# Same for station 104
print("104 Station total working time:", round((manager.station104 / 60) / 60, 2), "h", 
      "(", round((manager.station104 / manager.current_time) * 100, 2), "%)")

# Same for station 105
print("105 Station total working time:", round((manager.station105 / 60) / 60, 2), "h", 
      "(", round((manager.station105 / manager.current_time) * 100, 2), "%)")

# Prints the robot's total movement time in hours and its percentage compared to the total time
print("Robot total working time:", round(((manager.robotmovtime) / 60) / 60, 2), "h", 
      "(", round((manager.robotmovtime / manager.current_time) * 100, 2), "%)")