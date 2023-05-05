import re


# define regex pattern to extract information from subject line to a dictionary
sp = r'(?:RE:\s+|FW:\s+)?' \
     r'(?P<c_type>CofC|Certificate of conformance|CUSTOM CofC)\s+' \
     r'(?P<cert_number>\d+)\s+' \
     r'(?P<product_number>[\w-]+)\s+' \
     r'SO (?P<so_number>\w+)\s+' \
     r'LOT (?P<lot_number>[\w\.\-\d]+)\s+' \
     r'(?P<customer>.+)\s+' \
     r'(?P<c_number>\d+)\s+' \
     r'BP (?P<loc_number>[\d ]+)'

# compile for faster repeated use
subject_pattern = re.compile(sp)
