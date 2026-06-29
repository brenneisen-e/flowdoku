import sys, os, socket, time, uno
from com.sun.star.beans import PropertyValue

def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p

def wait_socket(port=2002, tries=240):
    for _ in range(tries):
        s=socket.socket()
        try:
            s.connect(('localhost',port)); s.close(); return True
        except Exception:
            s.close(); t=time.time()+0.5
            while time.time()<t: pass
    return False

def main():
    src=os.path.abspath(sys.argv[1]); outdir=sys.argv[2]; which=sys.argv[3] if len(sys.argv)>3 else "all"
    os.makedirs(outdir, exist_ok=True)
    assert wait_socket(), "UNO socket not ready"
    lc=uno.getComponentContext()
    resolver=lc.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", lc)
    ctx=resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    smgr=ctx.ServiceManager
    desktop=smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    url="file://"+src.replace("\\","/")
    doc=desktop.loadComponentFromURL(url,"_blank",0,(pv("Hidden",True),pv("ReadOnly",True),pv("FilterName","Impress MS PowerPoint 2007 XML")))
    exporter=smgr.createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", ctx)
    pages=doc.DrawPages
    idxs=range(pages.Count) if which=="all" else [int(x)-1 for x in which.split(",")]
    for i in idxs:
        page=pages.getByIndex(i)
        exporter.setSourceDocument(page)
        fdata=uno.Any("[]com.sun.star.beans.PropertyValue",(pv("PixelWidth",1600),pv("PixelHeight",900)))
        outurl="file://"+os.path.abspath(os.path.join(outdir,"s%02d.png"%(i+1))).replace("\\","/")
        exporter.filter((pv("MediaType","image/png"),pv("URL",outurl),fdata))
        print("rendered", i+1)
    doc.close(False)
main()
