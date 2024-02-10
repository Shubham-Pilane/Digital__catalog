import sys
import os
import subprocess
import easyocr
import pandas as pd
from pymongo import MongoClient
from gtts import gTTS
import vlc
import time
import speech_recognition as sr

# Global variable to store product name
global_product_name = ""

def speak(audioString):
    tts = gTTS(text=audioString, lang='en')
    filename = "voice.mp3"
    tts.save(filename)
    vlc.MediaPlayer(filename).play()
    time.sleep(1)

def capture_product_name():
    global global_product_name  # Access the global variable
    speak("Hi, what do you want to do?\n1. Add a single product\n2. Add bulk data\n3. Delete a product")
    choice = input("Enter your choice (1, 2, or 3): ")

    if choice == '1':
        global_product_name = input("Hello, what is your product name? ")
        inserted_document_id = add_single_product(global_product_name)
        return inserted_document_id
    
    elif choice == '2':
        excel_file_path = input("Enter the path to the Excel file: ")
        bulk_data(excel_file_path)
        sys.exit()  # Exit the program after adding bulk data
    
    elif choice == '3':
        speak("Please say the product name you want to delete.")
        delete_product_with_voice_command()
        sys.exit()  # Exit the program after deleting a product
    
    else:
        speak("Invalid choice. Please enter 1, 2, or 3.")
        return None

def add_single_product(product_name):
    connection_string = 'mongodb://localhost:27017'
    client = MongoClient(connection_string)
    db = client['Digital_catalog']
    collection = db['label']

    existing_document = collection.find_one({'product_name': product_name})

    if existing_document:
        print(f"Product '{product_name}' already exists with ID: {existing_document['_id']}")
        return existing_document['_id']
    else:
        document = {'product_name': product_name}
        result = collection.insert_one(document)
        inserted_id = result.inserted_id
        print(f"Inserted document ID: {inserted_id}")
        client.close()
        return inserted_id  

def bulk_data(excel_file_path):
    try:
        df = pd.read_excel(excel_file_path)
        connection_string = 'mongodb://localhost:27017'
        client = MongoClient(connection_string)
        db = client['Digital_catalog']
        collection = db['label']
        data = df.to_dict(orient='records')
        result = collection.insert_many(data)
        for inserted_id in result.inserted_ids:
            print(f"Inserted document ID: {inserted_id}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        client.close()

def delete_product(product_name):
    # Connect to MongoDB
    connection_string = 'mongodb://localhost:27017'
    client = MongoClient(connection_string)
    db = client['Digital_catalog']
    collection = db['label']

    # Delete the document with the specified product name
    result = collection.delete_one({'product_name': product_name})

    if result.deleted_count > 0:
        speak(f"Product '{product_name}' has been deleted.")
    else:
        speak(f"Product '{product_name}' not found.")

    # Close MongoDB client connection
    client.close()

def delete_product_with_voice_command():
    recognizer = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)

    try:
        product_name = recognizer.recognize_google(audio).capitalize()
        print(f"User said: {product_name}")
        delete_product(product_name)
    except sr.UnknownValueError:
        speak("Sorry, I couldn't understand your voice.")
    except sr.RequestError as e:
        speak(f"Error with the speech recognition service: {e}")

def process_image(image_path, document_id):
    reader = easyocr.Reader(['en'])
    result = reader.readtext(image_path, detail=0, paragraph=True)

    indexed_result = []
    for index, text in enumerate(result):
        if index == 0:
            words = text.split()
            indexed_words = list(enumerate(words))
            indexed_result.append((index, indexed_words))
        else:
            indexed_result.append((index, text))

    indexed_result_str = repr(indexed_result)
    insert_data(indexed_result_str, document_id)

def insert_data(indexed_result_str, document_id):
    connection_string = 'mongodb://localhost:27017'
    client = MongoClient(connection_string)
    db = client['Digital_catalog']
    collection = db['label']

    indexed_result = eval(indexed_result_str)
    data_dict = {'origin': '', 'size_US': '', 'size_UK': '', 'size_FR': '', 'size_JP': '', 'barcode': '', 'image': '', 'image_type': ''}

    if indexed_result and indexed_result[0][1]:
        for nested_index, word in indexed_result[0][1]:
            if nested_index == 2:
                data_dict['origin'] = word
            elif nested_index == 7:
                data_dict['size_US'] = float(word)
            elif nested_index == 8:
                data_dict['size_UK'] = float(word)
            elif nested_index == 9:
                data_dict['size_FR'] = float(word)
            elif nested_index == 10:
                data_dict['size_JP'] = float(word)

    if len(indexed_result) > 1:
        data_dict['barcode'] = indexed_result[1][1]

    collection.update_one({'_id': document_id}, {'$set': data_dict})
    print(f"Updated document ID: {document_id}")
    print("Updated data:")
    print(data_dict)
    client.close()

# ...

def update_image_info(document_id, folder_path):
    image_files = [f for f in os.listdir(folder_path) if f.endswith(('.jpg', '.jpeg', '.png'))]

    connection_string = 'mongodb://localhost:27017'
    client = MongoClient(connection_string)
    db = client['Digital_catalog']
    collection = db['label']

    global global_product_name  # Access the global variable

    for image_file in image_files:
        product_name, image_type = os.path.splitext(image_file)

        existing_document = collection.find_one({'product_name': product_name})
        if existing_document:
            collection.update_one(
                {'_id': existing_document['_id']},
                {'$set': {
                    'image': product_name,
                    'image_type': image_type[1:],
                }}
            )

            print(f"Updated document ID: {existing_document['_id']}")
            print("Updated data:")
            print(f"Image: {product_name}, Image Type: {image_type}")

    client.close()

# ...

if __name__ == "__main__":
    inserted_document_id = capture_product_name()

    print("Please scan the product label.")
    image_path = 'C:/Users/me/Desktop/label.jpg'  # Replace with the actual image path
    process_image(image_path, inserted_document_id)
    print("Product label scanned and values updated.")

    folder_path = 'C:/Users/me/Desktop/GK'  # Replace with the actual folder path
    print(f"Please add an image to the specified folder: {folder_path}")
    update_image_info(inserted_document_id, folder_path)
    print("Image information is updated.")
