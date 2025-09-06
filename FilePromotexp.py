
adress_email = ''
with open('C:\Mails.txt', 'w') as file:
            file.write(f"s&0&{adress_email}.pro@gmail.com&")
            
body = '''Un zombie (ou zombi) est, dans le folklore, une créature fictive souvent décrite comme un mort-vivant, ou une personne dont lesprit a été contrôlé par un pouvoir surnaturel ou une maladie. Dans la culture populaire, les zombies sont généralement représentés comme des créatures agressives, cherchant à se nourrir de chair humaine, et leur morsure peut transmettre une infection qui transforme les victimes en zombies. 
En termes plus généraux, le terme "zombie" peut également être utilisé pour désigner une personne qui semble dépourvue de volonté, d'émotion ou de conscience, agissant comme un automate. 
Voici quelques points clés à retenir sur les zombies :
Origines:
Le concept de zombie trouve ses origines dans le folklore haïtien, où il est associé au vaudou. '''       

with open('C:\EmailSend.txt', 'w') as file:
    file.write(f"{adress_email}@gmail.com&Evolution D'un mamifere&{body}&")
