from PIL import Image, ImageDraw, ImageFont
S=110; W,H=int(13.33*S),int(7.5*S)
DEEP=(0x04,0x6A,0x38); PALE=(0xF1,0xF6,0xE4); DARK=(0x1C,0x3D,0x26)
GREY=(0x75,0x78,0x7B); BODY=(0x3A,0x3A,0x3A); WHITE=(255,255,255); LIGHT=(0xEA,0xF3,0xDE)
def F(sz,b=False):
    return ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans%s.ttf"%("-Bold" if b else ""),int(sz*S/72.0))
def wrap(d,t,f,mw):
    out=[];cur=""
    for w in t.split():
        x=(cur+" "+w).strip()
        if d.textlength(x,font=f)<=mw: cur=x
        else: out.append(cur) if cur else None; cur=w
    if cur:out.append(cur)
    return out
def rr(d,x,y,w,h,fill,r=12): d.rounded_rectangle([x,y,x+w,y+h],radius=r*S/72.0,fill=fill)
def para(d,x,y,t,f,c,mw,lh=1.13):
    for ln in wrap(d,t,f,mw): d.text((x,y),ln,font=f,fill=c); y+=f.size*lh
    return y
def badge(img,d,x,y,dia,fill,icon):
    d.ellipse([x,y,x+dia,y+dia],fill=fill)
    ic=Image.open(icon).convert("RGBA"); isz=int(dia*0.52); ic=ic.resize((isz,isz))
    img.paste(ic,(int(x+(dia-isz)/2),int(y+(dia-isz)/2)),ic)

def base(title,sub):
    img=Image.new("RGB",(W,H),WHITE); d=ImageDraw.Draw(img)
    d.text((0.55*S,0.42*S),title,font=F(28,True),fill=(0x22,0x22,0x22))
    d.text((0.55*S,0.98*S),sub,font=F(16.5),fill=GREY)
    return img,d

# slide1
left0=0.55; total=13.33-1.1; GAP=0.32; CW=(total-3*GAP)/4
L=[left0+i*(CW+GAP) for i in range(4)]; CTOP=2.02; CH=4.45; BD=0.92; PX=0.30
img,d=base("DEX – Was seit April passiert ist","Eigener Plattform-Wechsel statt Entwicklung in Indien – mit messbarem Erfolg")
kpis=[("icons/refresh-cw.png","April","Neue Plattform live","Bewusst gegen die Entwicklung mit Indien entschieden (Komplexität, Skillset, Budget, Zeit). KI-gestützt eigenständig auf eine neue Architektur migriert – live & produktiv."),
("icons/calendar.png","29","Events koordiniert","Seit dem Wechsel 29 Events vollständig über DEX abgewickelt – Tendenz weiter steigend."),
("icons/users.png","~7.000","Teilnehmer","Rund 7.000 Anmeldungen gesteuert – inkl. vollautomatischer Mail- & Outlook-Anbindung, die den Assistenzen viel Zeit spart."),
("icons/check-circle.png","≈100%","Self-Service","Event-Erstellung läuft eigenständig durch die Assistenzen – ohne operative Einbindung von Nils & Eike. 72 Std. Eigenleistung gebucht.")]
for i,(ic,num,lab,desc) in enumerate(kpis):
    x=L[i]*S;y=CTOP*S;w=CW*S;h=CH*S; rr(d,x,y,w,h,PALE)
    badge(img,d,x+PX*S,y+0.34*S,BD*S,DEEP,ic)
    ny=y+0.34*S+BD*S+0.10*S; d.text((x+PX*S,ny),num,font=F(38,True),fill=DEEP)
    ly=ny+0.86*S; d.text((x+PX*S,ly),lab,font=F(15,True),fill=DARK)
    para(d,x+PX*S,ly+0.44*S,desc,F(11),BODY,(CW-2*PX)*S)
img.save("preview_1.png")

# slides 2/3
GX=0.30;GY=0.30;CW2=(total-GX)/2;CTOP2=1.92;CH2=((6.92-CTOP2)-GY)/2
COLX=[0.55,0.55+CW2+GX];ROWY=[CTOP2,CTOP2+CH2+GY];BD2=0.95;LPAD=0.32
def grid(title,sub,cards,fn):
    img,d=base(title,sub)
    for pos,c in enumerate(cards):
        r,co=divmod(pos,2);x=COLX[co]*S;y=ROWY[r]*S;w=CW2*S;h=CH2*S
        hl=c.get("hl"); rr(d,x,y,w,h,DEEP if hl else PALE)
        by=y+(h-BD2*S)/2; badge(img,d,x+LPAD*S,by,BD2*S,(WHITE if hl else DEEP),c.get("ia") if hl else c["icon"])
        tx=x+(LPAD+BD2+0.28)*S; tw=w-(tx-x)-0.30*S
        hcol=WHITE if hl else DARK; dcol=LIGHT if hl else BODY
        # estimate block height for vertical centering
        hl_lines=wrap(d,c["head"],F(16.5,True),tw); de_lines=wrap(d,c["desc"],F(12),tw)
        bh=len(hl_lines)*16.5*S/72*1.13 + 6*S/72 + len(de_lines)*12*S/72*1.15
        yy=y+(h-bh)/2
        for ln in hl_lines: d.text((tx,yy),ln,font=F(16.5,True),fill=hcol); yy+=16.5*S/72*1.13
        yy+=6*S/72
        for ln in de_lines: d.text((tx,yy),ln,font=F(12),fill=dcol); yy+=12*S/72*1.15
    img.save(fn)

grid("Neue Governance & Verankerung","Von der App zum zentralen Deloitte-Teilnehmermanagement-Tool",[
 dict(icon="icons/shield.png",head="Betriebsrats-Freigabe",desc="Neue, erweiterte BR-Freigabe steht in den nächsten Tagen – die Voraussetzung für den breiten Rollout."),
 dict(icon="icons/user-check.png",head="Power-User & Multiplikatoren",desc="Assistenzen aus T&T abgeholt; Eva Wienkamp, Annette Stoffel und Ebru Genctürk als Power-User etabliert."),
 dict(icon="icons/share-2.png",head="Ausweitung auf Service Lines",desc="Für Sondercalls bei SRT sowie T&L (Audit & Tax) eingeladen – weitere Assistenz-Calls folgen."),
 dict(icon="icons/send.png",head="Intranet-Kommunikation",desc="Intranet-Artikel in Vorbereitung, um DEX für die breite Deloitte-Welt freizuschalten."),
],"preview_2.png")

grid("Plan nach vorne & Budgetbedarf","Veröffentlichung, Schulung und Klärung der Finanzierung für das kommende FY",[
 dict(icon="icons/globe.png",head="Veröffentlichung & Listung",desc="DEX offiziell veröffentlichen und in den Deloitte Tools listen – als zentrales Teilnehmermanagement-Tool sichtbar machen."),
 dict(icon="icons/book-open.png",head="Schulungen & Onboarding",desc="Zentrale Info- und Onboarding-Termine aufsetzen, um die breite Masse befähigt an die App heranzuführen."),
 dict(icon="icons/user-plus.png",head="Kleines Core-Team",desc="Festes Junior-Team für die Begleitung der Schulungs- und Bekanntmachungs-Maßnahmen aufbauen."),
 dict(icon="icons/dollar-sign.png",ia="icons/dollar-sign-green.png",hl=True,head="Budget-Klärung",desc="Verlängerung des Budget-Codes bzw. eigene Budgetposition fürs kommende FY – für Schulung, Betrieb & Routinearbeiten. Bedarf: ~5.000 EUR (FY25/26), ~10.000 EUR ab FY26/27. Bestehender Code: GASR8771. Unterlage für den OMP-Call."),
],"preview_3.png")
print("ok")
