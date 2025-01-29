from StepBase import StepBase
import tkinter as tk
from tkinter import ttk, messagebox

class StepTwo(StepBase):
    def __init__(self):
        super().__init__()
        self.step = 2

    def setup_widgets(self, parent_frame):
        # Create step label
        self.step_label = ttk.Label(parent_frame, 
                                  text="Step 2: Copy and Paste your Fusing Coordination Report", 
                                  font=("Arial", 14))
        self.step_label.pack(pady=10)

        # Create paste widgets
        self.create_paste_widgets(parent_frame)
        self.paste_label.config(text="Paste Fusing Coordination Data Here")
        self.paste_label.pack(pady=10)
        self.text_frame.pack(pady=10)

        # Add process and skip buttons
        self.process_btn = ttk.Button(parent_frame, 
                                    text="Parse Data and Move to Next Step",
                                    command=self.save_data)
        self.process_btn.pack(pady=10)

        self.skip_btn = ttk.Button(parent_frame, 
                                 text="Skip This Step",
                                 command=self.next_step)
        self.skip_btn.pack(pady=10)

    def process_file(self):
        pass

    def save_data(self):
        pasted_data = self.paste_text.get("1.0", tk.END).strip()
        if not pasted_data:
            messagebox.showerror("Error", "Please paste data into the text box.")
            return

        try:
            lines = pasted_data.split("\n")
            header = lines[0].split("\t")
            data = []
            
            for line in lines[1:]:
                fields = line.split("\t")
                sequence = fields[0][:4]
                existing = fields[2]
                data.append([sequence, existing])

            with open("extractFusingCoordination_newOrExistingFusing.txt", "w") as file:
                file.write("Sequence\tExisting\n")
                for row in data:
                    file.write("\t".join(row) + "\n")

            messagebox.showinfo("Success", f"Data from Step {self.step} saved successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse pasted data: {e}")  