import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d


class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()

        # create widget variable for user input
        self.distance_file = tk.StringVar()
        self.elevation_file = tk.StringVar()
        self.roughness_coeff = tk.StringVar()
        self.slope = tk.StringVar()
        self.min_water_level = tk.StringVar()
        self.max_water_level = tk.StringVar()

        # create widgets
        self.distance_label = tk.Label(self, text="Distance CSV file")
        self.distance_entry = tk.Entry(self, textvariable=self.distance_file, state='readonly')
        self.distance_button = tk.Button(self, text="Browse", command=self.load_distance)
        self.elevation_label = tk.Label(self, text="Elevation CSV file")
        self.elevation_entry = tk.Entry(self, textvariable=self.elevation_file, state='readonly')
        self.elevation_button = tk.Button(self, text="Browse", command=self.load_elevation)
        self.roughness_coeff_label = tk.Label(self, text="Roughness Coefficient")
        self.roughness_coeff_entry = tk.Entry(self, textvariable=self.roughness_coeff)
        self.slope_label = tk.Label(self, text="Slope")
        self.slope_entry = tk.Entry(self, textvariable=self.slope)
        self.min_water_level_label = tk.Label(self, text="Min Water Level")
        self.min_water_level_entry = tk.Entry(self, textvariable=self.min_water_level)
        self.max_water_level_label = tk.Label(self, text="Max Water Level")
        self.max_water_level_entry = tk.Entry(self, textvariable=self.max_water_level)
        self.submit_button = tk.Button(self, text="Submit", command=self.generate_table)

        # add widgets to grid
        self.distance_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.distance_entry.grid(row=0, column=1, padx=5, pady=5, sticky="we")
        self.distance_button.grid(row=0, column=2, padx=5, pady=5)

        self.elevation_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.elevation_entry.grid(row=1, column=1, padx=5, pady=5, sticky="we")
        self.elevation_button.grid(row=1, column=2, padx=5, pady=5)

        self.roughness_coeff_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.roughness_coeff_entry.grid(row=2, column=1, padx=5, pady=5, sticky="we")

        self.slope_label.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.slope_entry.grid(row=3, column=1, padx=5, pady=5, sticky="we")

        self.min_water_level_label.grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.min_water_level_entry.grid(row=4, column=1, padx=5, pady=5, sticky="we")

        self.max_water_level_label.grid(row=5, column=0, padx=5, pady=5, sticky="w")
        self.max_water_level_entry.grid(row=5, column=1, padx=5, pady=5, sticky="we")

        self.submit_button.grid(row=6, column=1, pady=10)

    def load_distance(self):
        """
        Opens File Dialog to load distance CSV file.
        """
        file_path = filedialog.askopenfilename()
        self.distance_file.set(file_path)

    def load_elevation(self):
        """
        Opens File Dialog to load elevation CSV file.
        """
        file_path = filedialog.askopenfilename()
        self.elevation_file.set(file_path)

    def generate_table(self):
        # load data from files
        distance_path = self.distance_file.get()
        elevation_path = self.elevation_file.get()
        roughness_coeff = float(self.roughness_coeff.get())
        slope = float(self.slope.get())
        min_water_level = float(self.min_water_level.get())
        max_water_level = float(self.max_water_level.get())

        if not distance_path or not elevation_path:
            tk.messagebox.showerror("Error", "Please select distance and elevation files.")
            return

        try:
            distance_data = pd.read_csv(distance_path)
            elevation_data = pd.read_csv(elevation_path)
            matched_data = pd.merge(distance_data, elevation_data, on='Distance')
            sorted_data = matched_data.sort_values(by='Distance')
            distance = sorted_data['Distance'].to_numpy()
            elevation = sorted_data['Elevation'].to_numpy()

            table = self.calculate_table(distance, elevation, roughness_coeff, slope, min_water_level, max_water_level)

            self.plot_curve(table)

        except Exception as e:
            tk.messagebox.showerror("Error", f"Unable to generate table: {e}")

    def calculate_table(self, distance, elevation, roughness_coeff, slope, min_water_level, max_water_level, step=0.1):
        hydraulic_radius = 0
        wetted_perimeter = 0
        flow_area = 0
        water_levels = np.arange(min_water_level, max_water_level + step, step)
        flow_rates = []
        wetted_perimeters = []
        flow_areas = []
        for water_level in water_levels:
            flow_rate = self.manning_flow_rate(distance, elevation, water_level, roughness_coeff, slope)
            flow_rates.append(flow_rate)
            wetted_perimeter = 0
            flow_area = 0
            min_elevation = min(elevation)
            interp_func = interp1d(elevation, distance, kind="linear", fill_value="extrapolate")
            prev_dist = interp_func(min_elevation)
            prev_elev = min_elevation

            for elev, dist in zip(elevation, distance):
                if elev > water_level:
                    break
                # calculate the length of tangent line at lower end of the polyline segment
                dx = abs(prev_dist - dist)
                dh = abs(prev_elev - elev)
                L = (dx ** 2 + dh ** 2) ** 0.5
                prev_dist = dist
                prev_elev = elev
                # calculate wetted perimeter and flow area
                wetted_perimeter += L
                if elev >= water_level - depth:
                    dy = abs(prev_elev - water_level + depth)
                    flow_area += (dx * dy)
                else:
                    flow_area += ((water_level - elev) * dx)

            hydraulic_radius = flow_area / wetted_perimeter
            velocity = (1 / roughness_coeff) * hydraulic_radius ** (2 / 3) * slope ** (1 / 2)

            flow_rate = velocity * flow_area
            wetted_perimeters.append(wetted_perimeter)
            flow_areas.append(flow_area)
            flow_rates.append(flow_rate)

        table = pd.DataFrame(
            {"water_level": water_levels, "flow_rate": flow_rates, "wetted_perimeter": wetted_perimeters,
             "flow_area": flow_areas})
        return table


    def manning_flow_rate(self, distance, elevation, water_level, roughness_coeff, slope):
        hydraulic_radius = 0
        wetted_perimeter = 0
        flow_area = 0
        min_elevation = min(elevation)
        interp_func = interp1d(elevation, distance, kind="linear", fill_value="extrapolate")
        try:
            left_distance = interp_func(water_level)
            right_distance = interp_func(water_level + 0.01)
        except ValueError:
            return 0
        width = abs(right_distance - left_distance)
        depth = water_level - min_elevation
        prev_dist = interp_func(min_elevation)
        prev_elev = min_elevation

        for elev, dist in zip(elevation, distance):
            if elev > water_level:
                break
            # calculate the length of tangent line at lower end of the polyline segment
            dx = abs(prev_dist - dist)
            dh = abs(prev_elev - elev)
            L = (dx ** 2 + dh ** 2) ** 0.5
            prev_dist = dist
            prev_elev = elev
            # calculate wetted perimeter and flow area
            wetted_perimeter += L
            if elev >= water_level - depth:
                dy = abs(prev_elev - water_level + depth)
                flow_area += (dx * dy)
            else:
                flow_area += ((water_level - elev) * dx)

        hydraulic_radius = flow_area / wetted_perimeter
        velocity = (1 / roughness_coeff) * hydraulic_radius ** (2 / 3) * slope ** (1 / 2)
        flow_rate = velocity * flow_area

        return flow_rate


    def plot_curve(self, table):
        fig, ax1 = plt.subplots()
        ax2 = ax1.twinx()

        ax1.plot(table["water_level"], table["flow_rate"], color="blue", label="Flow Rate")
        ax2.plot(table["water_level"], table["wetted_perimeter"], color="green", label="Wetted Perimeter")
        ax2.plot(table["water_level"], table["flow_area"], color="red", label="Flow Area")

        ax1.set_xlabel("Water Level (m)")
        ax1.set_ylabel("Flow Rate (m^3/s)")
        ax2.set_ylabel("Wetted Perimeter (m), Flow Area (m^2)")
        ax1.set_title("Rating Curve")

        ps = ax1.get_lines() + ax2.get_lines()
        labs = [p.get_label() for p in ps]
        ax1.legend(ps, labs, loc="upper left")

        plt.show()


root = tk.Tk()
app = App(master=root)
app.mainloop()


