import inputMsgs

def getMenuCmd():
    return input(inputMsgs.menuCmd).upper()

def getPhrases():
    phrases = []
    while(True):
        print( str.format('All phrases: {}', phrases) )
        phrase = input('Enter phrase to match. Enter blank phrase when done: ')
        if phrase != '':
            phrases.append(phrase)
        elif phrase == '':
            if len(phrases) == 0:
                print('Must enter atleast one phrase...')
            else:
                break
    return phrases

def getCost():
    while True:
        costInput = input("Enter cost to be written: ")
        try:
            return float(costInput)
        except ValueError:
            print("Invalid number. Try again.")

def getConfirmation(msg):
    while True:
        resp = input(msg)
        uppercasedResp = resp.upper()
        if uppercasedResp == 'Y' or uppercasedResp == 'N':
            return uppercasedResp
        else:
            print('Invalid command...Try again')