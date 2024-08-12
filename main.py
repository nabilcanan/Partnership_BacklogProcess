import tkinter as tk
import tkinter.ttk as ttk
import os
from neotech_backlog import merge_contract_and_backlog
from creation_backlog import merge_contract_and_backlog_creation
from neotech_query import new_function_Neotech
from creation_query import new_function_Creation


def gui_create(main_window):
    style = ttk.Style()
    main_window.configure(bg="white")
    # Adjusting button size and font size
    style.configure("TButton", font=("Roboto", 12, "bold"), width=30, height=2)
    style.map("TButton", foreground=[('active', 'white')], background=[('active', '#007BFF')])

    title_label = ttk.Label(main_window, text="Welcome Partnership Team!",
                            font=("Segoe UI", 24, "underline"), background="white", foreground="#103d81")
    title_label.pack(pady=10)  # Reduced padding

    description_label = ttk.Label(main_window,
                                  text="This tool allows you to run the Backlog Query, then select either \n"
                                       "your Neotech Contract or Creation contract \n and process columns for price increases",
                                  font=("Roboto", 14), background="white", anchor="center",
                                  justify="center")
    description_label.pack(pady=10)

    instructions_button = ttk.Button(main_window, text="Open Instructions", command=open_instructions)
    instructions_button.pack(pady=10)  # Added padding for the button

    run_queries_button = ttk.Button(main_window, text="Run Backlog Query for Neotech", command=new_function_Neotech, style="TButton")
    run_queries_button.pack(pady=10)

    process_button = ttk.Button(main_window, text="Neotech Backlog Process", command=merge_contract_and_backlog)
    process_button.pack(pady=20)  # Added padding for the button

    run_queries_button = ttk.Button(main_window, text="Run Backlog Query for Creation", command=new_function_Creation, style="TButton")
    run_queries_button.pack(pady=10)

    process_button_creation = ttk.Button(main_window, text="Creation Backlog Process", command=merge_contract_and_backlog_creation)
    process_button_creation.pack(pady=20)  # Added padding for the button


def open_instructions():
    # Path to the instructions PowerPoint file
    instructions_path = r"P:\Partnership_Python_Projects\Backlog Query Process Neotech and Creation\Backlog Process for Neotech and Creation.pptx"
    # Open the file with the default application
    os.startfile(instructions_path)


if __name__ == "__main__":
    root = tk.Tk()
    root.title('Backlog Query Process')
    root.geometry('600x500')
    gui_create(root)
    root.mainloop()
