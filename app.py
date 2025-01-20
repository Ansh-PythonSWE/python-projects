import pandas as pd

class Computer:
    def __init__(self, id, brand, model, status):
        self.id = id
        self.brand = brand
        self.model = model
        self.status = status

    def __str__(self):
        return f"ID: {self.id}, Brand: {self.brand}, Model: {self.model}, Status: {self.status}"

class ComputerManagementSystem:
    def __init__(self):
        self.computers = []

    def add_computer(self, id, brand, model, status):
        computer = Computer(id, brand, model, status)
        self.computers.append(computer)
        print(f"Computer {id} added successfully.")

    def remove_computer(self, id):
        self.computers = [comp for comp in self.computers if comp.id != id]
        print(f"Computer {id} removed successfully.")

    def update_computer_status(self, id, status):
        for comp in self.computers:
            if comp.id == id:
                comp.status = status
                print(f"Computer {id} status updated to {status}.")
                return
        print(f"Computer {id} not found.")

    def list_computers(self):
        for comp in self.computers:
            print(comp)

    def save_to_excel(cms, filename):
                data = [{
                    "ID": comp.id,
                    "Brand": comp.brand,
                    "Model": comp.model,
                    "Status": comp.status
                } for comp in cms.computers]
                
                df = pd.DataFrame(data)
                df.to_excel(filename, index=False)
                print(f"Data saved to {filename}")
 
 
    def load_from_excel(cms, filename):
        try:
            df = pd.read_excel(filename)
        except FileNotFoundError:
            print(f"File {filename} not found.")
            return
        for _, row in df.iterrows():
                    cms.add_computer(row["ID"], row["Brand"], row["Model"], row["Status"])
        print(f"Data loaded from {filename}")
cms = ComputerManagementSystem()
while True:
                print("\n1. Add Computer")
                print("2. Remove Computer")
                print("3. Update Computer Status")
                print("4. List Computers")
                print("5. Save to Excel")
                print("6. Load from Excel")
                print("7. Exit")

                choice = input("Enter your choice: ")

                if choice == '1':
                    id = int(input("Enter Computer ID: "))
                    brand = input("Enter Computer Brand: ")
                    model = input("Enter Computer Model: ")
                    status = input("Enter Computer Status: ")
                    cms.add_computer(id, brand, model, status)
                elif choice == '2':
                    id = int(input("Enter Computer ID to remove: "))
                    cms.remove_computer(id)
                elif choice == '3':
                    id = int(input("Enter Computer ID to update: "))
                    status = input("Enter new status: ")
                    cms.update_computer_status(id, status)
                elif choice == '4':
                    cms.list_computers()
                elif choice == '5':
                    filename = input("Enter filename to save to (e.g., computers.xlsx): ")
                    cms.save_to_excel(filename)
                elif choice == '6':
                    filename = input("Enter filename to load from (e.g., computers.xlsx): ")
                    cms.load_from_excel(filename)
                elif choice == '7':
                    break
                else:
                    print("Invalid choice. Please try again.")
