import tkinter as tk
import speech_recognition as sr
import threading
import re
import pandas as pd   

def filter_data(year=None, name=None, price=None):
    try:
        df = pd.read_excel("full_data.xlsx")

        applied_filters = []

        if year:
            df = df[df["Year"] == year]
            applied_filters.append(f"Year: {year}")
        if name:
            df = df[df["Name"].str.lower() == name.lower()]
            applied_filters.append(f"Name: {name.title()}")
        if price:
            df = df[df["Price"] == price]
            applied_filters.append(f"Price: {price}")

        if not df.empty:
            with pd.ExcelWriter("filtered_data.xlsx", engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, index=False, sheet_name='Data')
        else:
            output_box.insert(tk.END, f"\n⚠️ No data found for {', '.join(applied_filters)}")

    except Exception as e:
        output_box.insert(tk.END, f"\n❌ Error: {str(e)}")

def recognize_voice():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        status_label.config(text="🎤 Listening...")
        try:
            audio = recognizer.listen(source, timeout=5)
            status_label.config(text="🧠 Processing...")
            text = recognizer.recognize_google(audio)
            output_box.delete("1.0", tk.END)
            output_box.insert(tk.END, f"You said: {text}")
            status_label.config(text="✅ Done.")

            year = None
            price = None

            # Extract name(s) from known list
            detected_names = [person for person in possible_names if person in text.lower()]

            # Extract price (excluding year)
            numbers = re.findall(r"\b(\d{4,6})\b", text)
            if numbers:
                for num in numbers:
                    n = int(num)
                    if not year or n != year:
                        price = n
                        break

            # If comparing multiple names
            if len(detected_names) >= 2:
                df = pd.read_excel("full_data.xlsx")
                comp_df = df[df["Name"].str.lower().isin(detected_names)]

                if not comp_df.empty:
                    with pd.ExcelWriter("comparison_data.xlsx", engine='openpyxl', mode='w') as writer:
                        comp_df.to_excel(writer, index=False, sheet_name='Comparison')
                    output_box.insert(tk.END, f"\n📊 Comparison data saved for: {', '.join([n.title() for n in detected_names])}")
                else:
                    output_box.insert(tk.END, f"\n⚠️ No data found for comparison: {', '.join(detected_names)}")
                return

            # If only one name, apply normal filters
            name = detected_names[0] if detected_names else None

        except Exception as e:
            status_label.config(text=f"❌ Error: {str(e)}")


# GUI Setup
root = tk.Tk()
root.title("Voice-Controlled Dashboard")
root.geometry("500x300")

title = tk.Label(root, text="🎙️ Voice Controller", font=("Arial", 18))
title.pack(pady=10)

btn = tk.Button(root, text="Speak", command=start_listening, font=("Arial", 14))
btn.pack(pady=10)

output_box = tk.Text(root, height=5, width=50)
output_box.pack(pady=10)

status_label = tk.Label(root, text="", font=("Arial", 12))
status_label.pack()

root.mainloop()


