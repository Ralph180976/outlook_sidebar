import win32com.client

def inspect_categories():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("--- Outlook Categories ---")
        if namespace.Categories.Count == 0:
            print("No categories found.")
            return

        for cat in namespace.Categories:
            print(f"Name: {cat.Name}, Color Enum: {cat.Color}")
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect_categories()
