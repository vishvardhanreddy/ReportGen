import xml.etree.ElementTree as ET
from StringIO import StringIO
assets = {}

def find_rows(anp, parent_id):
    for row in anp.findall("shelf"):
        RUsize = (row.get('RUsize'))
        RUpos = (row.get('RUpos'))
        rackPos = (row.get('rackPos'))
        rackNum = (row.get('rackNum'))
        id = (row.get('id'))
        ty = (row.get('type'))
        
        assets[ty] = {'RUsize': RUsize,
                          'RUpos': RUpos,
                          'rackPos': rackPos,
                          'rackNum': rackNum,
                          'id': id,
                          'type': ty}
        
    for row in anp.findall("slot"):
        occupancy = (row.get('occupancy'))
        pos = (row.get('pos'))
                
        assets[ty] = {'occupancy': occupancy,
                           'pos': pos}
        
    child_anp = row.find("anp")
    if child_anp is not None:
            find_rows(child_anp, ty)
      
    print assets.values()
        
def main():
    
    tree = ET.parse('1.Collector//neupdate.xml')
    root = tree.getroot()            
    first_anp = root.find('.//anp')
    find_rows(first_anp, None)
    
    

if __name__ == '__main__':
        main()