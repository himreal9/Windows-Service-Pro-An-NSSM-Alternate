import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import json
import ctypes
import os
import datetime

class ServiceManagerApp:
    def __init__(self, root):
        if not self.is_admin():
            messagebox.showwarning("Warning", "This application requires administrative privileges to manage services. Please run as an administrator.")
            root.destroy()
            return
        self.root = root
        self.root.title("Service Manager")
        

        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Create tabs
        self.new_service_tab = ttk.Frame(self.notebook)
        self.available_service_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.new_service_tab, text="New Services")
        self.notebook.add(self.available_service_tab, text="Available Services")

        # Variables for entry widgets
        self.new_service_name_var = tk.StringVar()
        self.new_exe_location_var = tk.StringVar()

        # Entry widgets for New Services tab
        tk.Label(self.new_service_tab, text="Service Name:").grid(row=0, column=0, sticky="e")
        service_name_entry = tk.Entry(self.new_service_tab, textvariable=self.new_service_name_var)
        service_name_entry.grid(row=0, column=1, columnspan=2, sticky="we", pady=5)

        tk.Label(self.new_service_tab, text="Executable Location:").grid(row=1, column=0, sticky="e")
        exe_location_entry = tk.Entry(self.new_service_tab, textvariable=self.new_exe_location_var)
        exe_location_entry.grid(row=1, column=1, sticky="we", pady=5)
        
        browse_button = tk.Button(self.new_service_tab, text="Browse", command=self.browse_location)
        browse_button.grid(row=1, column=2, sticky="e", pady=5)

        start_button = tk.Button(self.new_service_tab, text="Create and Start Service", command=self.start_service)
        start_button.grid(row=2, column=0, columnspan=3, pady=10)

        # Treeview for Available Services tab
        self.available_services_tree = ttk.Treeview(self.available_service_tab, columns=("Service Name", "Process Name", "Status"))
        self.available_services_tree.heading("#0", text="Service Name")
        self.available_services_tree.heading("#1", text="Process Name")
        self.available_services_tree.heading("#2", text="Status")
        self.available_services_tree.column("#0", width=200, anchor="center")
        self.available_services_tree.column("#1", width=200, anchor="center")
        self.available_services_tree.column("#2", width=100, anchor="center")

        self.available_services_tree.grid(row=0, column=0, padx=10, pady=10)
        button_frame = tk.Frame(self.available_service_tab)
        button_frame.grid(row=1, column=0, pady=10)


        start_button_0 = tk.Button(button_frame, text="Start Service", command=self.start_service_again)
        start_button_0.grid(row=0, column=0, padx=5)

        stop_button = tk.Button(button_frame, text="Stop Service", command=self.stop_service)
        stop_button.grid(row=0, column=1, padx=5)

        delete_button = tk.Button(button_frame, text="Delete Service", command=self.delete_service)
        delete_button.grid(row=0, column=2, padx=5)


        

        # Load available services from JSON
        self.load_available_services()
    def is_admin(self):
        try:
            # Check if the script is running with administrative privileges
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
        
    def browse_location(self):
        file_path = filedialog.askopenfilename()
        self.new_exe_location_var.set(file_path)

    def start_service(self):
        service_name = self.new_service_name_var.get()
        exe_location = self.new_exe_location_var.get()

        if not service_name or not exe_location:
            messagebox.showerror("Error", "Please enter both service name and executable location.")
            return

        try:
            project_location = os.getcwd()
            exe_name = 'Service_Generator.exe'
            exe_location0 = os.path.join(project_location, exe_name)
            timestamp = datetime.datetime.now().strftime("%H%M%S")
            exe_location=exe_location.replace('/','//')
            new_exe_location = exe_location.replace('.exe','_'+timestamp+'.exe')
            new_exe_name=new_exe_location.split('//')[-1]
            os.rename(exe_location, new_exe_location)
            create_command = f'sc create {service_name} binPath= "{exe_location0} {new_exe_location}" start= auto'
            subprocess.run(create_command)
            subprocess.run(["sc", "start", service_name])
            data = {service_name: {"exe_name": new_exe_name, "status": "Running"}}
            json_file_path = os.path.join(project_location, "process.json")
            try:
                with open(json_file_path, 'r') as json_file:
                    existing_data = json.load(json_file)
            except FileNotFoundError:
                existing_data = {}
            existing_data.update(data)
            with open(json_file_path, 'w') as json_file:
                json.dump(existing_data, json_file, indent=4)

            messagebox.showinfo("Success", "Service created and started successfully.")

            # Update available services list
            self.load_available_services()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to create and start service. Error: {str(e)}")

    def load_available_services(self):
        try:
            import os
            project_location = os.getcwd()
            json_file_path = os.path.join(project_location, "process.json")
            try:
                with open(json_file_path, 'r') as json_file:
                    existing_data = json.load(json_file)
            except FileNotFoundError:
                existing_data = {}

            with open(json_file_path, 'w') as json_file:
                json.dump(existing_data, json_file, indent=4)
            # Load services from JSON
            with open(json_file_path, 'r') as json_file:
                data = json.load(json_file)

            # Clear existing items in the treeview
            for item in self.available_services_tree.get_children():
                self.available_services_tree.delete(item)

            # Populate treeview with available services
            for service_name, info in data.items():
                self.available_services_tree.insert("", "end", text=service_name, values=(info["exe_name"], info["status"]))

        except FileNotFoundError:
            # If the file does not exist, create an empty JSON
            with open(json_file_path, 'w') as json_file:
                json.dump({}, json_file)

    def stop_service(self):
        selected_item = self.available_services_tree.selection()

        if not selected_item:
            messagebox.showerror("Error", "Please select a service to stop.")
            return
        elif self.available_services_tree.item(selected_item, 'values')[1]!="Running":
            messagebox.showerror("Error", "Please select a running service to stop.")
            return

        service_name = self.available_services_tree.item(selected_item, 'text')
        exe_name = self.available_services_tree.item(selected_item, 'values')[0]    
        
        try:
            # Terminate the service subprocess using psutil
            try:
                subprocess.run(['taskkill', '/F', '/IM', exe_name], check=True)
            except:
                pass
            subprocess.run(["sc", "stop", service_name])

            try:
                subprocess.run(['taskkill', '/F', '/IM', exe_name], check=True)
            except:
                pass

            # Remove the service from JSON
            with open('process.json', 'r') as json_file:
                data = json.load(json_file)

            if service_name in data:
                data[service_name]["status"]="Stopped"

            with open('process.json', 'w') as json_file:
                json.dump(data, json_file)

            messagebox.showinfo("Success", "Service stopped successfully.")

            # Update available services list
            self.load_available_services()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to stop service. Error: {str(e)}")


    def delete_service(self):
        selected_item = self.available_services_tree.selection()

        if not selected_item:
            messagebox.showerror("Error", "Please select a service to delete.")
            return

        service_name = self.available_services_tree.item(selected_item, 'text')
        exe_name = self.available_services_tree.item(selected_item, 'values')[0]
        
        try:
            # Terminate the service subprocess using psutil
            try:
                subprocess.run(['taskkill', '/F', '/IM', exe_name], check=True)
            except:
                pass
            subprocess.run(["sc", "stop", service_name])
            subprocess.run(["sc", "delete", service_name])

            try:
                subprocess.run(['taskkill', '/F', '/IM', exe_name], check=True)
            except:
                pass

            # Remove the service from JSON
            with open('process.json', 'r') as json_file:
                data = json.load(json_file)

            if service_name in data:
                del data[service_name]

            with open('process.json', 'w') as json_file:
                json.dump(data, json_file)

            messagebox.showinfo("Success", "Service deleted successfully.")

            # Update available services list
            self.load_available_services()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to stop service. Error: {str(e)}")

    def start_service_again(self):
        selected_item = self.available_services_tree.selection()

        if not selected_item:
            messagebox.showerror("Error", "Please select a service to start.")
            return
        elif self.available_services_tree.item(selected_item, 'values')[1]!="Stopped":
            messagebox.showerror("Error", "Please select a stopped service to run.")
            return

        service_name = self.available_services_tree.item(selected_item, 'text')
        exe_name = self.available_services_tree.item(selected_item, 'values')[0]

        if not exe_name:
            messagebox.showwarning("Warning", "No process associated with the selected service.")
            return
        
        try:

            subprocess.run(["sc", "start", service_name])

            with open('process.json', 'r') as json_file:
                data = json.load(json_file)

            if service_name in data:
                data[service_name]["status"]="Running"

            with open('process.json', 'w') as json_file:
                json.dump(data, json_file)

            messagebox.showinfo("Success", "Service started successfully.")

            # Update available services list
            self.load_available_services()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to start service. Error: {str(e)}")
        

if __name__ == "__main__":
    root = tk.Tk()
    app = ServiceManagerApp(root)
    root.mainloop()