#print string on a new line use /n
print("girrafe\nacademy")
# If you want to print a quotation mark without python interpreting it as the end of a string use backslash
print("Giraffe\"academy")
#This is how to print a backslash only
print("giraffe/academy")
#typing a string variable
phrase = "Giraffe Academy"
print (phrase + "is such a cool forum")
#function
#funciton is a special block of operation that is used to perform a specific operation
#converting a string to lower case
print (phrase.lower())
#converting a string to upper case
print(phrase.upper())
#check if a string is entirely upper or lower
print(phrase.isupper())
print(phrase.islower())

#use functions in combination.
print(phrase.upper().isupper())

#To find the length of a string
print(len(phrase))
#To print a specific character in a string
print(phrase[0])
#The index function will tell us where a specific character or string is located.
print(phrase.index("G"))
print(phrase.index("i"))
print(phrase.index("Acad"))
print(phrase.index("demy"))
#if you to input a character that is not in the string, it will throw an error.
#replace function.
phrase = "mongoose Academy"
print(phrase)
print(phrase.replace("mongoose", "elephant"))
#Working with numbers
print(3)
print(2.93939)
print(-2.939)
print(3+4)
print(10/4)
print(3*2+9)
print(3*(2+9))
#modulus
print(10 % 3)
print(17% 13)
print(-17%12)
#store numbers in variable containers
my_num=(17%11)
print(my_num)

#convert numbet to string.
print(str(my_num))
print(str(my_num) + " my lucky number")
#python will not allow you to print a combination of numbers and strings without first converting them to string.

#absolute of a number
my_num= (-17)
print(abs(my_num))
#power of a number
print(pow(4,3))
# max or min function function
print(max(100, 1999))
print(min(100,200,19))
#rounding off
print(round(3.7))

#importing python math to access some more mathematical functions.
from math import *
#floor or ceiling of a number
print(floor(3.8))
print(ceil(4.3))
print(sqrt(54))
# how to get input from a user
name = input("Please enter your name")
age = input("Please enter your age")
print("hello " + name + " you are " + age + " years old")

# how to make a very basic calculator

num_1 = input("enter the first number")
num_2 = input("enter the second number")
result = num_2 + num_1
print(result)

#this gives a wrong result because by default, when you get an input from a user, python will conver it to a string.
#so you have to convert these strings to numbers eg using int function

num_1 = input("enter the first number")
num_2 = input("enter the second number")
result = int(num_2) + int(num_1)
print(result)
# the int function will not alow to add decimal numbers together and hence you have to use float

num_1 = input("enter the first number")
num_2 = input("enter the second number")
result = float(num_2) + float(num_1)
print(result)

#mad libs game
color= input("enter you favorite color")
plural_noun= input("enter a plural noun")
celebrity=input("enter your favorite celebrity")

print("roses are" + color)
print(plural_noun + " are blue")
print("i love "+ celebrity)

#list in python
#you can even store strings, numbers or boolens in a list
friends = ["jackson","wahome", "muthui","muchunu",false, 2 ]
print(friends)

#list in python
#you can even store strings, numbers or boolens in a list
# you can access the whole list or the individual elements based on their indices
friends = ["jackson","wahome", "muthui","muchunu",2]
print(friends)
print(friends[2])
print(friends[0])
print(friends[4])
# you can also access the individual elements based on the indices from the right using negatives
print(friends[-1])
#portions of the list
print(friends[1:3])
#note that it only grabs the elements up to but not including 3
#another way is to type 1: and will return all the elemnts after index 1
print(friends[1:])
#modifying elements in a list
friends[1]="honorable"
print((friends[1]))



#list functions
lucky_numbers=[1,16,27, 24 ,57,94,100]
friends=["john","albert","kevin", "james", "wanjiru", "joan", "wanyonyi"]
#lists are invaluable in python
print(friends)
#the extend function will allow you to append a list onto another list
friends.extend(lucky_numbers)
print(friends )
#add individual elements into a list
friends.append("mogaka")
print(friends)
#add an item in the middle of the list
friends.insert(1,"muraya")
print(friends)
#remove an element
friends.remove("joan")
print(friends)
#getting rid of the last element in the list
friends.pop()
print(friends)

#check if a certain value is in the list
print(friends.index("james"))
print(friends.index("wanjiru"))
# If you check for a value which in not in the list python will throw an error
#count the number of similar elements in the list
friends.append("kevin")
print(friends)
print(friends.count("kevin"))
#sort a list in alphabetical order
friends=["john","albert","kevin", "james", "wanjiru", "joan", "wanyonyi"]
print(friends.sort())
print(friends)
lucky_numbers=[111,166,27, 24 ,57,94,100]
lucky_numbers.sort()
print(lucky_numbers)
#reversing the order in a list
lucky_numbers.reverse()
print(lucky_numbers)
#copying a list
mabeshte= friends.copy()
print(mabeshte)

#remove all the elements from the list
friends.clear()
print(friends)

#tuples
#is a type of a data structure. it is a container where we can store different values it is very similar to a list
# eg we can store cordinates inside a tuple
coordinates =(4,5)
#tuples are immutable: they can't be changed or modified.
#if you try to modify python will throw an error.
print(coordinates[0])
print(coordinates[1])
#you can create a list of tuples
coordinates=[(4,5), (9,10), (15,17)]
#by a rich majority, lists will be used more often unless you wonna store data that won't be changed.






#functions
#a bunch of code that is used to perform a specific funtion. they allow you to organize you code much better
# create a function that says high to the user

#used def followed by parentheses and a colon to define a function
#all code in the function should be indented
#then you need to call the function to execute it. codes inside a fuction will not be executed by default
# function names need to be in all lowercase and different names separated by an underscore
def say_hi():
    print("hello jackson")
say_hi()

#order of execution
def say_hi_to_me():
    print("hello jackson")
print("mambo mzae")
say_hi_to_me() #this is how to call a function
print("uko fiti?")


#Parameters are pieces of information which are given to the function
def gotea(name):
    print("hello " + name)
gotea("jackson")
gotea("Wahome")

def biodata(name,age):
    print("hello "+ name + "you are "+ age)
biodata("jackson ", "90 yrs")
biodata("wahome ","65 yrs")

def personal_data(name,age):
    print("hello "+ name + "you are "+ str(age) + " years old")
personal_data("jackson ", 90)
personal_data("wahome" ,69)

#you can parse any type of data inside a function eg boolean, arrays, strings, integers
#generally it's  good idea to break your code into different functions.



#return statement
# make a function which can cube a number

def cube(num):
    return num*num*num
print(cube(900000000))

#the return keyword is going to return the value 27 when the function is called.
def cube(num):
    return num*num*num
result =cube(4)
print(result)

#the return statement is very useful in getting information from  a statement.
# the return statement breaks the function and hence not other line after it will be executed




#if statements
is_male = True
if is_male:
    print("you are male")
#But in this case, if is_male = false, it wil not print anything
#anything that come after the if statement and has the indentation, will be executed.
# we can also use another keyword which is else.

is_male = False
if is_male:
    print("you are male")
else:
    print("you are female")

#using or $ and key words
is_male = True
is_tall= False
if is_male and is_tall:
    print("you are male who is very tall")
else:
    print("aai ww ni bure kabisa")

is_male = True
is_tall = False
if is_male or is_tall:
    print("you are male or who it tall or both")
else:
    print("aai ww ni bure kabisa")

    # using if else
is_male = False
is_tall = False
if is_male and is_tall:
    print("you are male who is tall")
elif is_male and not is_tall:
    print("you are male who it not tall. unaona kama utapata mrembo?")
else:
    print("aai ww ni bure kabisa. Tafuta kamba ujinyonge magoti")





#if statements with comparisons
def maxnum (num1, num2, num3):
    if num1 >=num2 and num1 >= num3:
        return num1
    elif num2 >= num1 and num2>= num3:
        return num2
    else:
        return num3


print(maxnum(200, 205, 190))
# you can compare numbers, strings, booleans etc
# equal to ==
# not equal to !=




# building a better calculator with all basic arithmetic operations.
num1= float(input("enter the first number"))
op= input("enter the operator e.g")
num2= float(input("enter the second number"))

if op == "+":
    print(num1 + num2)
elif op == "-":
    print(num1 - num2)
elif op == "*":
    print(num1 * num2)
else:
    print(num1/num2)

    # using dictionaries in python
    # dictionary is a special structure in python which allows us to store key-value pairs
    # just like in a normal dictionary a word would have a difinition, in this context, they word would be the key and the value
    # would be the actual definition.
    # program to convert the three-month name to full-month name
    # all keys have to be unique
    month_conversions = {
        "jan": "january",
        "feb": "february",
        "mar": "march",
        "apr": "april",
        "may": "may",
        "jun": "june",
        "jul": "july",
        "aug": "august",
        "sep": "september",
        "oct": "october",
        "nov": "november",
        "dec": "december",
    }
    print(month_conversions["dec"])
    print(month_conversions.get("dec"))
    # the get function will help return a default value when key is not mappable to any value
    print(month_conversions.get("decc", "not a valid key"))
    # the keys don't have to be strings. They can also be numbers.




# while loop
# it allows us to loop through and execute a block of code multiple times
x = 1
while x < 15:
    print(x)
    x += 1
print("The loop is finished")  # remember to remove indent here
