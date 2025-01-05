def format_amt(amt):
    amt=amt.replace(',','')
    negative=False
    if '-' in amt:
            negative=True
            amt=amt.replace('-', '')
    if len(amt)>5:
        x=amt.index('.')
        newamt=amt[x-3:x+3]
        counter=1
        # print(amt[:x-3:])
        for i in amt[:x-3:][::-1]:
            counter=counter+1
            if counter==2:
                newamt=i+','+newamt
                counter=0
            else:
                newamt=i+newamt
        if negative==True:   
            return '-'+newamt
        else:
             return newamt
    else:
        if negative==True:
             return '-'+amt
        else:
            return amt
    
print(format_amt('428665.89'))