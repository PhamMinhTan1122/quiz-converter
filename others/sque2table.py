from docx import Document

def split_sequence(sequence):
    # Split the input sequence into numbers and characters
    numbers = []
    characters = []
    for item in sequence:
        number, character = ''.join(filter(str.isdigit, item)), ''.join(filter(str.isalpha, item))
        numbers.append(number)
        characters.append(character)
    return numbers, characters

def convert_sequence_to_table(numbers, characters):
    document = Document()
    table = document.add_table(rows=1, cols=10)
    table.style = 'Table Grid'

    for i in range(0, len(numbers), 10):
        row_cells = table.add_row().cells
        for j in range(10):
            index = i + j
            if index < len(numbers):
                row_cells[j].text = f"{numbers[index]} - {characters[index]}"

    document.save('output.docx')

if __name__ == "__main__":
    input_sequence = "1D 2A 3A 4D 5B 6D 7D 8C 9A 10C 11A 12A 13B 14B 15A 16B 17A 18B 19C 20A 21D 22D 23B 24C 25D 26A 27C 28A 29A 30A 31D 32B 33A 34A 35A 36A 37A 38D 39A 40C 41A 42C 43C 44C 45A 46A 47B 48C 49A 50A 51D 52C 53D 54C 55A 56A 57A 58C 59D 60A 61C 62C 63A 64A 65D 66A 67A 68C 69B 70A 71C 72C 73C 74D 75A 76A 77D 78A 79A 80D 81A 82B 83D 84C 85B 86D 87D 88A 89D 90D 91 92B 93D 94A 95B 96A 97B 98C 99D 100B 101C 102C 103D 104C 105A 106D 107B 108C 109D 110A 111B 112B 113D 114B 115B 116A 117B 118A 119D 120C 121B 122A 123B 124C 125B 126C 127D 128C 129A 130D 131C 132B 133A 134D 135B 136C 137A 138C 139A 140B 141D 142A 143B 144B 145B 146C 147C 148D 149D 150C 151D 152A 153B 154C 155A 156B 157A 158C 159A 160C 161A 162D 163B 164B 165C 166D 167A 168A 169B 170A 171C 172D 173A 174C 175B 176A 177B 178A 179C 180D 181A 182C 183B 184C 185A 186C 187D 188B 189B 190D 191A 192C 193D 194A 195B 196B 197A 198B 199C 200A 201B 202C 203C 204D 205C 206B 207C 208A 209A 210A 211B 212A 213D 214D 215D 216D 217D 218C 219C 220A 221C 222B 223C 224A 225B 226A 227A 228A 229A 230A 231A 232A 233A 234D 235B 236A 237C 238B 239C 240B 241A 242B 243B 244A 245A 246A 247C 248D 249C 250C 251C 252B 253B 254C 255C 256B 257A 258B 259B 260A 261C 262C 263C 264C 265A 266D 267C 268B 269B 270A 271B 272A 273C 274B 275D 276C 277B 278A 279A 280A 281A 282A 283A 284A 285C 286B 287B 288A 289A 290A 291B 292A 293A 294A 295A 296B 297A 298A 299B 300C"
    sequence_list = input_sequence.split()

    numbers, characters = split_sequence(sequence_list)
    convert_sequence_to_table(numbers, characters)
