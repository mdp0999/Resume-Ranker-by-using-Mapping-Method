file = open("outfile.txt","r").read().split("\n")
file = [x for x in file if x != '']
dict_list = []
for file in file:
    file = file.split("\t")
    file1 = float(file[1].replace("%",""))
    dict1 = {"name":file[0].replace(".docx",""), "score":file1}
    dict_list.append(dict1)
    
from operator import itemgetter
newlist = sorted(dict_list, key=itemgetter('score'), reverse=True)
write_file = open("main_results.txt", "w")
a = 0
def get_values_as_tuple(dict_list, keys):
    return [tuple(d[k] for k in keys) for d in dict_list]
new = get_values_as_tuple(newlist, ['name', 'score'])
for s in new:
    a += 1
    write_file.write("Rank: "+str(a)+"\t"+"name: "+str(s[0])+"\t"+"Overall Percentage: "+str(s[1])+"\n")
    