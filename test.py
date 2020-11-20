def mail_html():
    s_order = []
    img_address = []
    point = 0
    a = 0
    f = open("index.txt",'rt',encoding='UTF8')
    data = f.read()
    while a != -1:
        a = data.find('<img',point)
        s_order.append(a)
        point = a+1
    del s_order[-1]

    for i in s_order:
        target_contents = ''
        count = i+10
        s = data[count]
        while s != '"':
            target_contents = target_contents + s
            count = count+1
            s = data[count]
        img_address.append(target_contents)
    #print(img_address)
    f.close()
    
    fw = open("index.txt",'w')
    for i in img_address:
        
    

mail_html()
