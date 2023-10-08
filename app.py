import constants
import utilities
import in_out_vob as iov


def make_your_choice():
    print("##################################################")
    print("##      Select an option:                        #")
    print("###      1.Report IN_OUT and VOB               ##")
    print("##################################################")

    while True:
        choice = input("Please enter you choice: ")

        if choice == "1":
            iov.run_scripts()
            break
        else:
            print("Unavailable choice. Please enter a valid choice")


if __name__ == '__main__':
    try:
        make_your_choice()
    except Exception as e:
        print(f"Si è verificato un errore: {str(e)}")

