from datetime import datetime, timedelta
import winsound
import pandas as pd

def make_sound():
    duration = 1500
    freq = 600
    winsound.Beep(freq, duration)

def export_to_excel(tasks, filename="tasks.xlsx"):
    if not tasks:
        print("No tasks to export.")
        return

    data = []
    for task_name, details in tasks.items():
        row = {
            "Task Name": task_name,
            "Description": details['description'],
            "Start Date": details['start_date'],
            "Deadline": details['deadline_date'],
            "Status": details['status'],
            "Priority": details['priority']
        }
        data.append(row)

    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Tasks exported successfully to {filename}")

def add(tasks):
    new_task = input("Enter task name: ")
    if new_task in tasks:
        print("Task already exists!")
    else:
        new_desc = input("Enter task description: ")
        new_start_date = input("Enter start date & time (YYYY-MM-DD HH:MM AM/PM): ")
        new_deadline = input("Enter deadline date & time (YYYY-MM-DD HH:MM AM/PM): ")
        priority = input("Enter priority (High/Medium/Low): ").capitalize()
        
        tasks[new_task] = {
            "description": new_desc,
            "start_date": new_start_date,
            "deadline_date": new_deadline,
            "status": "Pending",
            "priority": priority
        }
        print(f"Task '{new_task}'is added")

def update(tasks):
    update_task = input("Enter the task name you want to update: ")
    if update_task in tasks:
        ch = input("Is your task completed? (yes/no): ").strip().lower()
        if ch == 'yes':
            tasks[update_task]['status'] = "Completed"
            print(f"Congratulations! Task '{update_task}' is completed.")
        else:
            new_desc = input("Enter new description: ")
            new_start_date = input("Enter new start date & time (YYYY-MM-DD HH:MM AM/PM): ")
            new_deadline = input("Enter new deadline date & time (YYYY-MM-DD HH:MM AM/PM): ")
            priority = input("Enter new priority (High/Medium/Low): ").capitalize()

            tasks[update_task] = {
                "description": new_desc,
                "start_date": new_start_date,
                "deadline_date": new_deadline,
                "status": "Pending",
                "priority": priority
            }
            print(f"Task '{update_task}' is updated.")
    else:
        print("Task not found.")

def delete_tasks(tasks):
    delete_task = input("Enter task to delete: ")
    if delete_task in tasks:
        del tasks[delete_task]
        print(f"Task '{delete_task}' deleted.")
    else:
        print("Task not found.")

def display_tasks(tasks):
    if not tasks:
        print("\nNo tasks available.\n")
        return

    print("\n" + "-" * 135)
    print(f"{'Task Name'.ljust(15)} | {'Priority'.ljust(8)} | {'Description'.ljust(25)} | {'Start Date & Time'.ljust(20)} | {'Deadline'.ljust(20)} | {'Time Left'.ljust(20)} | {'Status'.ljust(10)}")
    print("-" * 135)

    for task, details in tasks.items():
        deadline = datetime.strptime(details['deadline_date'], "%Y-%m-%d %I:%M %p")
        time_left = deadline - datetime.now()

        if 0 < time_left.total_seconds() <= 300:
            print(f" WARNING: Task '{task}' has only 5 minutes left!")
            make_sound()

        if time_left.total_seconds() > 0:
            days, seconds = divmod(time_left.total_seconds(), 86400)
            hours, seconds = divmod(seconds, 3600)
            minutes, seconds = divmod(seconds, 60)
            time_left_str = f"{int(days)}d {int(hours)}h {int(minutes)}m {int(seconds)}s"
        else:
            time_left_str = "Deadline Passed"

        print(f"{task.ljust(15)} | {details['priority'].ljust(8)} | {details['description'].ljust(25)} | {details['start_date'].ljust(20)} | {details['deadline_date'].ljust(20)} | {time_left_str.ljust(20)} | {details['status'].ljust(10)}")

    print("-" * 135)

def task_manager():
    tasks = {}
    print("----WELCOME TO THE TASK MANAGEMENT APP----")

    total_task = int(input("Enter how many tasks you want to add = "))
    for i in range(1, total_task + 1):
        task_name = input(f"Enter task {i} name: ")
        task_desc = input(f"Enter description for '{task_name}': ")
        start_date = input(f"Enter start date & time for '{task_name}' (YYYY-MM-DD HH:MM AM/PM): ")
        deadline_date = input(f"Enter deadline date & time for '{task_name}' (YYYY-MM-DD HH:MM AM/PM): ")
        priority = input(f"Enter priority for '{task_name}' (High/Medium/Low): ").capitalize()
        
        tasks[task_name] = {
            "description": task_desc,
            "start_date": start_date,
            "deadline_date": deadline_date,
            "status": "Pending",
            "priority": priority
        }

    display_tasks(tasks)

    while True:
        try:
            operation = int(input("\nEnter an option:\n1-Add\n2-Update\n3-Delete\n4-View\n5-Export to Excel\n6-Exit\n"))

            if operation == 1:
                add(tasks)
            elif operation == 2:
                update(tasks)
            elif operation == 3:
                delete_tasks(tasks)
            elif operation == 4:
                display_tasks(tasks)
            elif operation == 5:
                export_to_excel(tasks)
            elif operation == 6:
                print("Thanks for using Task Manager...")
                break
            else:
                print("Invalid Input. Please enter a number from 1 to 6.")
        except ValueError:
            print("Invalid Input. Please enter a valid number.")

task_manager()