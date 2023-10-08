import constants
import utilities
import in_out_vob as iov
import reportPdlCheckPdl as pdlcheck


def make_your_choice():
    print("##################################################")
    print("##      Select an option:                        #")
    print("###      1.Report IN_OUT and VOB               ##")
    print("##################################################")

    while True:
        choice = input("Please enter you choice: ")

        if choice == "1":
            iov.run_scripts_report_in_out_pob()
            break
        elif choice == "2":
            pdlcheck.run_scripts_report_pdl_check()
            break
        else:
            print("Unavailable choice. Please enter a valid choice")


if __name__ == '__main__':
    try:
        make_your_choice()
    except Exception as e:
        print(f"Si è verificato un errore: {str(e)}")

