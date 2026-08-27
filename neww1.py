'''from google.genai import Client

GOOGLE_API_KEY="AIzaSyCDIt0Q_bzLDrplLuWhsryu9dbvLtUDJoU"

client = Client(api_key=GOOGLE_API_KEY)
response = client.models.generate_content(
    model="gemini-2.5-flash",
    contents="Explain AI in simple words"
)

print(response.text)

from google.genai import Client

GOOGLE_API_KEY="AIzaSyCDIt0Q_bzLDrplLuWhsryu9dbvLtUDJoU"

client = Client(api_key=GOOGLE_API_KEY)

while True:
    user_input = input("You: ")

    if user_input.lower() == "exit":
        break

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=user_input
    )

    print("Bot:", response.text)
    '''
import fastapi