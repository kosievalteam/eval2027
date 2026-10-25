"""HWP 5.0 (binary) -> table cells JSON. Usage: python3 -I hwp_tables.py in.hwp out.json"""
import sys, zlib, struct, json, olefile

def records(data):
    i=0
    while i+4<=len(data):
        h=struct.unpack('<I',data[i:i+4])[0]; i+=4
        tag=h&0x3ff; level=(h>>10)&0x3ff; size=(h>>20)&0xfff
        if size==0xfff:
            size=struct.unpack('<I',data[i:i+4])[0]; i+=4
        yield tag, level, data[i:i+size]; i+=size

def para_text(body):
    s=[]; j=0
    while j+2<=len(body):
        ch=struct.unpack('<H',body[j:j+2])[0]; j+=2
        if ch in (1,2,3,11,12,14,15,16,17,18,21,22,23): j+=14          # extended ctrl (16 bytes)
        elif ch==9: s.append('\t'); j+=14
        elif ch in (10,13): s.append('\n')
        elif ch<32: pass
        else: s.append(chr(ch))
    return ''.join(s)

def main(path, out):
    ole=olefile.OleFileIO(path)
    hdr=ole.openstream('FileHeader').read()
    compressed=struct.unpack('<I',hdr[36:40])[0]&1
    secs=sorted([e for e in ole.listdir() if e[0]=='BodyText'], key=lambda e:int(e[1][7:]))
    tables=[]; cur=None; cell=None; tbl_level=None
    for e in secs:
        raw=ole.openstream(e).read()
        data=zlib.decompress(raw,-15) if compressed else raw
        for tag,level,body in records(data):
            if tag==71 and body[:4][::-1]==b'tbl ':
                cur={'cells':[]}; tables.append(cur); tbl_level=level; cell=None
            elif tag==77 and cur is not None and cur.get('rows') is None:
                cur['rows'],cur['cols']=struct.unpack('<HH',body[4:8])
            elif tag==72 and cur is not None and level==tbl_level+1:
                # LIST_HEADER (cell): nparas(2) attr(4) [pad?] col row colspan rowspan
                nparas,attr=struct.unpack('<hI',body[:6])
                col,row,cs,rs=struct.unpack('<HHHH',body[8:16])
                cell={'row':row,'col':col,'rowspan':rs,'colspan':cs,'text':[]}
                cur['cells'].append(cell)
            elif tag==67 and cell is not None and level==tbl_level+2:
                cell['text'].append(para_text(body))
    for t in tables:
        for c in t['cells']: c['text']='\n'.join(c['text']).strip()
    json.dump(tables,open(out,'w',encoding='utf-8'),ensure_ascii=False,indent=1)
    for k,t in enumerate(tables):
        print(k, t.get('rows'), t.get('cols'), len(t['cells']), repr(t['cells'][0]['text'][:40]) if t['cells'] else '')
main(sys.argv[1], sys.argv[2])
