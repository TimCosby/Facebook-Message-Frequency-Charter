from os.path import isfile
from source.messenger_complier import get_messages
from source.messenger_complier import create_file
from source.excel_modifier import options

def generate(answer='recreate'):
    if 'recreate' in answer:
        message_count = {}

        while True:
            try:
                get_messages(message_count)
                break
            except:
                print('Retrying...')

        create_file(message_count)

if __name__ == '__main__':
    if isfile('database.xlsx'):
        # If existing file exists
        generate(answer=input('Recreate database or use existing? (Recreate/existing):\n').lower())
    else:
        generate()

    options()
