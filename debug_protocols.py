import sys
import webbrowser
import os

print("Testing ms-todo URL...")
try:
    os.startfile("ms-todo://")
    print("ms-todo:// started successfully")
except Exception as e:
    print("ms-todo failed:", e)

print("Testing outlook URL...")
try:
    os.startfile("outlookcal:")
    print("outlookcal: started successfully")
except Exception as e:
    print("outlookcal failed:", e)
