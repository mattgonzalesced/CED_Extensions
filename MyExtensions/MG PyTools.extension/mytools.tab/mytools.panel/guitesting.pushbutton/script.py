from pyrevit import forms

# Main execution
if __name__ == "__main__":
    try:
        # Display a simple hello message
        forms.alert("Hello", title="Greetings")
    except Exception as e:
        # Handle errors gracefully
        forms.alert("An error occurred: " + str(e), title="Error")