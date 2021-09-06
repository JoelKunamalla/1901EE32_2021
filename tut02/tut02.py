def comp(input) :
    i=0
    yu = False
    iy =[]
    while( i < len(input)):
        y= input[i]
        if(type(y)!= int):
            yu = True
            iy.append(input[i])
        i=i+1
    if(yu == True):
        print("Please enter a valid input list.  Invalid inputs detected :" , iy)
          
    else :
        az = []
        c=0
        i=0
        while(i < len(input)) :
            cv= False
            for x in az :
                if(x == input[i] ):
                    c= c+1
                    cv=True
            if(cv==False):        
             az.append(input[i])        
            if(len(az) > 5):
             az.pop(0)
           
            i= i+1
        print( "score:" , c)
input = [1,4, 3, 5, 8,9,9]
comp(input)