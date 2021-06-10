#TODO: Create a letter using starting_letter.txt 
#for each name in invited_names.txt
#Replace the [name] placeholder with the actual name.
#Save the letters in the folder "ReadyToSend".
    
#Hint1: This method will help you: https://www.w3schools.com/python/ref_file_readlines.asp
    #Hint2: This method will also help you: https://www.w3schools.com/python/ref_string_replace.asp
        #Hint3: THis method will help you: https://www.w3schools.com/python/ref_string_strip.asp

#TODO: Test with outlook, pdf etc.
#TODO: Create simple GUI
#TODO: Add options for output
#TODO: Split generate_letters() into more functions, make it more clear to read

####################### CHANGE SETTINGS HERE ######################

NAMES_LIST = "100 Days of Code\\221_Mail Merge\\Mail Merge Project Start\\Input\\Names\\invited_names.txt"
STARTING_LETTER = "100 Days of Code\\221_Mail Merge\\Mail Merge Project Start\\Input\\Letters\\starting_letter.txt"
OUTPUT_FOLDER = "100 Days of Code\\221_Mail Merge\\Mail Merge Project Start\\Output\\"

####################### DO NOT TOUCH THIS ######################


class MailMerge():


    def invited_names_list(self):
        with open (NAMES_LIST, "r") as names:
            names = names.readlines()
            return names


    def generate_letters(self, names_list):
        for name in names_list:
            name = name.strip()
            with open(STARTING_LETTER, "r+") as letter:
                letter = letter.read()
                personal_letter = letter.replace("[name]", name)
                with open(f"{OUTPUT_FOLDER}{name}.txt", "w") as file:
                    file.write(personal_letter)


Robot = MailMerge()
Robot.generate_letters(Robot.invited_names_list())
