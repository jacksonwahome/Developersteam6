
#How to swap variables
a=5
b=6
a=b
b=a
print(a)
print(b)
# in this case the program prints the values a and b as 6
# you need to put the value a in a temp variable
a=5
b=6
temp = a
a=b
b= temp
print(a)
print(b)
# is there an easier way of doing it without using a third variable?
a=5 #101 you require 3 bits for 5
 b=6 #110 you also require 3 bits for 5
a=a+b #1011 but you require 4 bits to store 11
 b=a-b
a=a-b
 print(a)
 print(b)

# you can decide to use a caret symbol which does not consume extra memory
a=5
 b=6
a=a^b
 b=a^b
a=a^b
 print(a)
 print(b)

 or you can use commas
a,b=b,a
print(a)
print(b)


# accessing previous command in IDLE
# you need to navigage to options/configure idle/and change the key to one of your liking (eg up key)
#its still not possible to scroll on idle
# also how do you access it in command prompt
x=2+4+9
x=2+4+9

#bitwise Operators (complement, And, Or, xor,left shift, right shift)
#Python 3.10.4 (tags/v3.10.4:9d38120, Mar 23 2022, 23:13:41) [MSC v.1929 64 bit (AMD64)] on win32
#Type "help", "copyright", "credits" or "license()" for more information.
# accessing previous command in IDLE
# also how do you access it in command prompt
x=2+4+9
x=2+4+9
#bitwise Operators (complement, And, Or, xor,left shift, right shift)
#complement (~) tilde
~12
-13
# complement of a number is the reverse format  in binary
# for instance the reverse of 1 is 0
bin(12)
'0b1100'
bin(13)
'0b1101'
# no that was wrong
# no that was wrong
# we do not store negative numbers in the system
# we need to convert it first to a positive number
# 2s complement is equal to 1 complement + 1 complement
# binary of 13 00001101
#1s complement is 11110010
#2s complement =  +1= 11110011
# 11110011 is the two's complement of 13 which is -13
# the complement of 12 is -13



# Bitwise and
# and is for logical operators
# Bitwise and is denoted as &
12&13
12


# convert 12 into binary format which is 00001100
# convert 13 into binary format which is 00001101
# taking 1 as true and 0 as false then   00001100 ( this is the result) which 12 after you put correspond 1 or zero below them)


# Bitwise or
# Bitwise or is denoted as |
12|13
13
#00001100
#00001101
#00001101 (since 1 and zero give 1 two 1s give 1 and two zeros give 0

# Trying another bitwise and
25&31
25
# sometimes you can get a different number
25&30
24
# xor operator looks at odd number of 1 and gives a 1 and even number of 1 and gives a 0
# 00 gives 0
#01 gives 1
#10 gives 1
# 11 gives 0
25&30

# bitwise xor is denoted by ^ symbol
12^13
# here it means you convert the two numbers to binary and then whenever you  have same numbers,
# you put a 0 and whenever you have different number you put a 1
#00001100
#00001101
#00000001 ( and the answer is 1
25^30 #( ans is 7)
# left shift
10<<2
#1010
# 1010 is the same as 1010.0000
# it means you shift two bits to the left and the result is
# 101000.00 the ans is 40

# right shift
10>>2 # in right shift it means that you are losing bits unlike left shift
#10.____ and the ans is 2

#Importing math functions
x=sqrt(25)
# this evaluation gives you an error. You have to import math function to use them
import math
x=math.sqrt(25)
# well this is not working in pycharm but it is working in idle. ( to check later)
math.sqrt(25)
# still not working
# working with Pycharm Video 17





