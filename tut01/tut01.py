
def comp(list1) :
    i=0 
    j=0 
    k=0
    while(i <  len(list1)) :
        y = list1[i]
        x = y%10
        y=y//10
        y1= True   
        while(y >0) : 
            if(abs(x - y%10)==1) :
                y1 = True
            else :
                j= j+1
                print(  " NO  -" , list1[i]  , "not a meraki number")
                y1 = False
                break
            x= y%10
            y = y//10
        if(y1 == True ) :
            k = k+1
            print( "Yes  -" , list1[i] , "is a  meraki number ")
        i = i +1    
        
        
        
    print("The  input list contains" , k , "meraki  numbers  and " , j , " non meraki numbers " )
            
list1 =[12, 14, 56, 1]
comp(list1)

