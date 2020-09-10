goblinNames = ["Greg", "Jim"]
goblinMoney = [5, 10]

response = input("Which goblin would you like to view ")
print(goblinNames[int(response)] + " has " + str(goblinMoney[int(response)]) + " gold coins.")