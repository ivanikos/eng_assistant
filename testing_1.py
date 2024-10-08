


a = 'P800010 - LATER - in the support list - SD-530'
b = 'P800140 - SD-539 - correct'
c = ' - LATER - need to double check LOAD TAG'

print(a.replace("- in the support list", "").replace(" - ", "&").split("&"))