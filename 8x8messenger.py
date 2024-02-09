import pandas as pd
import pyperclip
import pyautogui
import keyboard
import sys
import time

phone_number = 'phone number'

#Image paths
image_path= (r"C:\Users\David Huynh\Documents\Work\Programs\February_Deal.JPG")


# specify the file path and sheet name

file_path = r"C:\Users\David Huynh\Documents\Work\Exports\Working\[WORK] Customer Phone Book.xlsx"
sheet_name = "1"

# read the Excel file into a DataFrame
df = pd.read_excel(file_path, sheet_name)

# specify the column name
column_name = "Cell Phone"

#remove NaN
df = df.dropna(subset=['Cell Phone'])
# remove duplicates
df = df.drop_duplicates(subset=['Cell Phone'])

# read the values from the specified column
column_values = df[column_name].values

count = 1

# loop through the values and copy each one to the clipboard
for index, value in enumerate(column_values):
    
    if df.loc[index, 'Checked'] == 'read':
        continue
    else:
    
        isStop = False
        
        pyperclip.copy(str(value))
        copied_text = pyperclip.paste()
        print(f'{count}: {value} copied to clipboard')
        df.loc[index, 'Checked'] = 'read'

        #8x8 Automation
        pyautogui.click(302,490)
        pyautogui.click(302,415)
        pyautogui.click(340,129)
        
        #last message
        #17706321700
        pyautogui.typewrite(phone_number)#copied_text)
        pyautogui.sleep(1)
        
        #Click on Contact Icon
        pyautogui.click(243,225)
        pyautogui.sleep(1)
        #Click on user message
        
        pyautogui.doubleClick(403,451)
        pyautogui.hotkey('ctrl','c')
        
        if "STOP".lower() in pyperclip.paste().lower():
            print("The word 'STOP' is in the clipboard text.")
            isStop = True
        else:
            print("The word 'STOP' is not in the clipboard text.")
            

        pyautogui.doubleClick(401,415)
        pyautogui.hotkey('ctrl','c')

        if "STOP".lower() in pyperclip.paste().lower():
            print("The word 'STOP' is in the clipboard text.")
            isStop = True
        else:
            print("The word 'STOP' is not in the clipboard text.")
        
        # The mouse triple clicks to try and "grab" and copy the text
        # at the current location
        
        pyautogui.tripleClick(414, 433)
        pyautogui.hotkey('ctrl','c')

        # Checks if message in clipboard is already present in the the program.
        if "1.".lower() in pyperclip.paste().lower():
            print("Message already sent.")
            isStop = True
            
        pyautogui.tripleClick(414, 363)
        pyautogui.hotkey('ctrl','c')

        if "1.".lower() in pyperclip.paste().lower():
            print("Message already sent.")
            isStop = True
            
            
        if isStop == False:
            pyautogui.click(457,519)
            pyautogui.typewrite("Valentine's is around the corner and we're here to show you some love!")
            pyautogui.hotkey('shift','enter')
            pyautogui.typewrite('Beautiful and loving deals just for you.')
            pyautogui.hotkey('shift','enter')
            
            pyautogui.write(" ")
            pyautogui.hotkey('shift','enter')
            
            pyautogui.typewrite('Available now at Nailsbestbuy.com.')
            pyautogui.hotkey('shift','enter')
            pyautogui.typewrite("#Nailsbestbuy")
            pyautogui.hotkey('shift','enter')
            
            pyautogui.write(" ")
            pyautogui.hotkey('shift','enter')
            
            pyperclip.copy('Text/Nhắn: (337) 317-5717')
            pyautogui.hotkey('ctrl','v')
            pyautogui.hotkey('shift','enter')
            pyperclip.copy('Call/Gọi: (337) 991-2477')
            pyautogui.hotkey('ctrl','v')
            pyautogui.hotkey('shift','enter')
            pyautogui.typewrite('Website: nailsbestbuy.com')
            pyautogui.hotkey('shift','enter')
            pyautogui.typewrite('Facebook: https://www.facebook.com/search/top?q=inail%20supply%2C%20llc')
            
            pyautogui.hotkey('shift','enter')
            pyautogui.write(" ")
            pyautogui.hotkey('shift','enter')
            
            pyautogui.typewrite('1. Please reply STOP if you would like to stop receiving deals.')
            
            # click on the attach button
            pyautogui.click(398,524)

            # wait for the file picker dialog box to appear
            pyautogui.sleep(2)

            # type the file path of the image
            pyautogui.typewrite(image_path)

            # press enter to attach the image
            pyautogui.press("enter")
            
            pyautogui.sleep(4)
            
            # Send the image
            pyautogui.click(760,453)
            pyautogui.click(760,322)

            
            # Save the dataframe to the same excel

            new_file_path = r"C:\Users\David Huynh\Documents\Work\Exports\Working\[NEW] Customer Phone Book.xlsx"
            df.to_excel(new_file_path, sheet_name = sheet_name, index = False)
            
            count+=1
            if keyboard.is_pressed('q'):  # if key 'q' is pressed 
                print('Program stopped')
                sys.exit()  # stop the program
            print(' ')

#print the count
print(f'{count} rows processed')
