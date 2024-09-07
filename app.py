import random
import xlwt
import xlrd
from xlutils.copy import copy  # To append data to an existing Excel file
from flask import Flask, request, jsonify, send_from_directory, session
from flask_session import Session
from dateutil import parser
from datetime import datetime
import razorpay
import os

app = Flask(__name__)

# Razorpay Configuration
razorpay_client = razorpay.Client(auth=("", ""))

# Configure session management
app.config['SECRET_KEY'] = 'your_secret_key'
app.config['SESSION_TYPE'] = 'filesystem'  # Store session data in the filesystem
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True
Session(app)

# Path to store the Excel file
EXCEL_FILE = 'bookings.xls'

# Function to create the Excel file if it doesn't exist
def create_excel_file():
    if not os.path.exists(EXCEL_FILE):
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('Bookings')

        # Add headers
        headers = ["Sl no", "Customer name", "No of tickets", "Date", "Time"]
        for col_num, header in enumerate(headers):
            sheet.write(0, col_num, header)

        # Save the file
        workbook.save(EXCEL_FILE)

# Create the Excel file with headers if it doesn't exist
create_excel_file()

# Random greeting responses to add variation
greeting_responses = [
    "Hello! Ready to book your tickets?",
    "Hi! Let's get your tickets booked.",
    "Hey! Let me help you with your ticket booking.",
    "Good day! Ready to reserve your tickets?"
]

# Helper function to get the response based on the state
def get_response(user_message):
    user_message = user_message.lower()  # Convert message to lowercase for consistency
    state = session.get('state', 'initial')  # Get the state, default to 'initial'
    
    print(f"Current state: {state}, User message: {user_message}")  # Debugging output

    
    # Always prioritize 'restart' and 'help' commands
    if "restart" in user_message:
        session.clear()  # Reset the session when restarting
        session['state'] = 'initial'
        return "No problem! Let's start the booking process over. Please say 'book' to begin."

    if "help" in user_message:
        return ("I can help you book tickets. Here's what you can do:\n"
                "- Say 'book' to start booking tickets.\n"
                "- Provide the date in DD-MM-YYYY format and time in HH:MM format.\n"
                "- Let me know how many tickets you need and your name.\n"
                "- Type 'restart' if you want to start over.\n"
                "How can I assist you now?")

    # Always prioritize greetings and reset the conversation state
    if any(greeting in user_message for greeting in ["hello", "hi", "hey", "good morning", "good afternoon", "good evening"]):
        session.clear()  # Clear session on greeting to start fresh
        session['state'] = 'initial'  # Reset the state on greeting
        return random.choice(greeting_responses)

    # Initial state: starting the booking process
    if any(bookMessage in user_message for bookMessage in ["book", "buy", "visit"]):
        session['state'] = 'collecting_date'  # Set the state to 'collecting_date'
        return "Great! Please provide the date you'd like to visit (e.g., DD-MM-YYYY)."
    
    if "cancel" in user_message:
        session['state'] = 'initial'
        return "No worries, cancelled."

    # Collecting date
    elif state == 'collecting_date':
        try:
        # Try parsing the date using dateutil parser to accept flexible formats
            user_date = parser.parse(user_message).date()
            today = datetime.today().date()

            if user_date < today:
                return "It looks like you've entered a past date. Please provide a date in the future (e.g., DD-MM-YYYY)."

            session['date'] = user_message
            session['state'] = 'collecting_time'
            return "Great! Now, what time would you like to visit (e.g., HH:MM)?"
        except ValueError:
            return "The date format seems incorrect. Please use DD-MM-YYYY."

    # Collecting time
    elif state == 'collecting_time':
        try:
            # Parsing time with AM/PM format
            datetime.strptime(user_message, "%I:%M %p")
            session['time'] = user_message
            session['state'] = 'collecting_tickets'
            return "Perfect! How many tickets would you like to book?"
        except ValueError:
            return "The time format seems incorrect. Please use the format HH:MM AM/PM."


    # Collecting number of tickets
    elif state == 'collecting_tickets':
        if user_message.isdigit():
            session['tickets'] = int(user_message)
            session['state'] = 'collecting_name'
            return "Got it. Can I have your name to complete the booking?"
        else:
            return "Please enter a valid number of tickets."

    # Collecting user's name
    elif state == 'collecting_name':
        session['name'] = user_message
        session['state'] = 'confirmation'
        return (f"Thank you, {user_message}! You've booked {session['tickets']} ticket(s) for {session['date']} at {session['time']}. "
                "Does everything look correct? Please type 'yes' to confirm or 'no' to restart the booking.")

    # Confirmation
    elif state == 'confirmation':
        if "yes" in user_message:
            # Proceed to payment
            session['state'] = 'awaiting_payment'
            return "Proceeding to payment. Click the 'Pay Now' button below to complete your payment."

        else:
            session['state'] = 'initial'
            return "Let's start over. How can I assist you today?"
    else :
        return "Sorry, I couldn't Understand!"

# Function to save booking details to Excel file
def save_booking_to_excel():
    # Open the existing Excel file
    rb = xlrd.open_workbook(EXCEL_FILE)
    sheet = rb.sheet_by_index(0)
    next_row = sheet.nrows  # The next available row

    # Copy the content of the existing file to allow appending
    wb = copy(rb)
    writable_sheet = wb.get_sheet(0)

    # Add the booking details
    writable_sheet.write(next_row, 0, next_row)  # Serial number
    writable_sheet.write(next_row, 1, session['name'])  # Customer name
    writable_sheet.write(next_row, 2, session['tickets'])  # Number of tickets
    writable_sheet.write(next_row, 3, session['date'])  # Date
    writable_sheet.write(next_row, 4, session['time'])  # Time

    # Save the updated file
    wb.save(EXCEL_FILE)

@app.route('/')
def index():
    return send_from_directory('', 'bot.html')

# Create Razorpay order for the amount
@app.route('/create_order', methods=['POST'])
def create_order():
    amount = session.get('tickets', 1) * 100 * 100  # Ticket price is ₹100 in paise
    order = razorpay_client.order.create(dict(
        amount=amount,
        currency='INR',
        payment_capture=1  # Automatic capture
    ))
    return jsonify(order)

# Handle payment verification
@app.route('/payment_verification', methods=['POST'])
def payment_verification():
    data = request.json
    try:
        # Verify the payment signature
        razorpay_client.utility.verify_payment_signature({
            'razorpay_order_id': data['order_id'],
            'razorpay_payment_id': data['payment_id'],
            'razorpay_signature': data['signature']
        })
        
        # On successful payment, save the booking details
        save_booking_to_excel()
        
        return jsonify({'status': 'Payment successful!'}), 200
    except razorpay.errors.SignatureVerificationError:
        return jsonify({'status': 'Payment verification failed!'}), 400

@app.route('/chat', methods=['POST'])
def chat():
    try:
        user_message = request.json.get('message')
        if not user_message:
            return jsonify({"response": "I didn't understand that. Please try again."}), 400

        bot_response = get_response(user_message)

        # Debugging output
        print(f"Session state: {session.get('state')}, User message: {user_message}")

        return jsonify({"response": bot_response})
    except Exception as e:
        print(f"Error: {str(e)}")  # Log any errors
        return jsonify({"response": "Something went wrong. Please try again later."}), 500

if __name__ == '__main__':
    app.run(debug=True)


