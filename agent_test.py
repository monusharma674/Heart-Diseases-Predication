from google.genai import Client

# Create Gemini Client
GOOGLE_API_KEY="AIzaSyCDIt0Q_bzLDrplLuWhsryu9dbvLtUDJoU




" \
""
client = Client(api_key=GOOGLE_API_KEY)

# Tool: Cost Calculator
def calculate_cost(issue):
    if "gas" in issue.lower():
        return "Estimated Cost: ₹2500 - ₹3500"
    elif "cooling" in issue.lower():
        return "Estimated Cost: ₹1500 - ₹2500"
    elif "noise" in issue.lower():
        return "Estimated Cost: ₹800 - ₹1500"
    else:
        return "Estimated Cost: ₹500 - ₹2000"

# Agent Loop
while True:
    user_input = input("Customer: ")

    if user_input.lower() == "exit":
        break

    # AI Diagnosis
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=f"""
        You are an expert AC technician.
        Diagnose the issue and suggest solution.
        Problem: {user_input}
        """
    )

    diagnosis = response.text

    # Agent Decision (Tool Calling)
    cost = calculate_cost(user_input)

    print("\n🤖 Technician AI:")
    print(diagnosis)
    print(cost)
    print("If problem persists, technician visit recommended.\n")