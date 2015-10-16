import uno
from com.sun.star.awt import Size
from com.sun.star.awt import Point
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt.FontSlant import ITALIC
from collections import OrderedDict
from sqlalchemy import create_engine, func
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from songstodb import Song, Verse

engine = create_engine('sqlite:////home/juho/Dropbox/srk/laulut.db', echo=False)
Base = declarative_base()

#soffice --accept'=socket,host=localhost,port=2002;urp;StarOffice.Service'

class Presentation:

    def __init__(self):
        localContext = uno.getComponentContext()
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
        desk = ctx.ServiceManager.createInstanceWithContext( "com.sun.star.frame.Desktop", ctx)
        self.document = desk.getCurrentComponent()

    def NewPageAtEnd(self):
        """Create new page at the end"""
        self.pagecount = self.document.DrawPages.Count
        self.curpage = self.document.DrawPages.insertNewByIndex(self.pagecount)

class Slide:

    def __init__(self, pres, is_wide=False):
        pres.NewPageAtEnd()
        pres.curpage.Layout = 1
        pres.curpage.remove(pres.curpage.getByIndex(0))
        self.drawpage = pres.curpage
        self.body = pres.curpage.getByIndex(0)
        self.text = self.body.getText()
        self.cursor = self.text.createTextCursor()
        self.location = Point()
        self.size = Size()
        self.location.X = 700
        self.location.Y = 700
        if is_wide:
            self.size.Width = self.body.Size.Width-1000
        else:
            self.size.Width = self.body.Size.Width-8000
        self.size.Height = self.body.Size.Height
        self.body.setPosition(self.location)
        self.body.setSize(self.size)

    def InsertHeader(self, headertext, color, italics = False):
        color = CodeColor(color)
        #self.cursor.ParaAdjust = 3	
        self.cursor.CharColor = color
        self.cursor.CharHeight=26
        if italics:
            self.cursor.CharPosture = ITALIC
        else:
            self.cursor.CharPosture = NONE
        self.text.insertString(self.cursor, headertext + '\n',False) 	

    def InsertSectionTitle(self, current):
        sections = ['Johdanto','Sana','Ylistys ja rukous','Ehtoollinen','Lähettäminen']
        self.cursor.CharColor = CodeColor('white')
        self.cursor.CharHeight=11
        for num, sec in enumerate(sections):
            if sec == current:
                self.cursor.CharColor = CodeColor('yellow')
            self.cursor.CharPosture = ITALIC
            self.cursor.CharPosture = NONE
            if num < len(sections)-1:
                sec = sec + ' | '
            self.text.insertString(self.cursor, sec, False) 	
            self.cursor.CharColor = CodeColor('white')


    def UpdateToc(self, currentsect,currenttitle):
        sections = OrderedDict()
        sections['Johdanto'] = ['Alkulaulu', 'Alkujuonto ja seurakuntalaisen sana']
        sections['Sana'] = ['päivän laulu','evankeliumi','saarna','synnintunnustus ja synninpäästö','uskontunnustus']
        sections['Ylistys ja rukous'] = ['Ylistys- ja rukouslauluja','Esirukous']
        sections['Ehtoollinen'] = ['Pyhä','Ehtoollisrukous','Kolehti ja ehtoollisen jako']
        sections['Siunaus ja lähettäminen'] = ['Herran siunaus','Loppujuonto','Loppulaulu']

        self.text.insertString(self.cursor, '\n'*2, False) 	
        self.cursor.CharHeight=20

        for title in sections[currentsect]:
            if title == currenttitle:
                title = ' --> ' + title
            else:
                title = '       ' + title
            self.text.insertString(self.cursor, title + '\n',False) 	
            self.cursor.CharColor = CodeColor('white')


def CodeColor(color):
    if color == 'white':
        return 16777215
    elif color == 'yellow':
        return 13421568

def LuoOtsikkodia(ThisComponent, yellowtext, whitetext):
    #Create a new slide at the end
    ThisComponent.DrawPages.insertNewByIndex(ThisComponent.DrawPages.Count)


Session = sessionmaker(bind=engine)
session = Session()



laulu = session.query(Song)

subquery = session.query(Song.id).filter(func.lower(Song.title) == 'lauluista kaunein').subquery()
query = session.query(Verse.content).filter(Verse.song_id.in_(subquery))
sakeistot = query.all()

messu = Presentation()
#
#sivu = Slide(messu,True)
#sivu.InsertSectionTitle('Johdanto')
#sivu.UpdateToc('Johdanto','Alkulaulu')
#
sivu = Slide(messu,True)
#sivu.InsertHeader(sakeistot[0][0],'white')
sakt = '\n' + sakeistot[0][0]
sivu.InsertHeader(sakt,'white')
#sivu.InsertHeader('Valossa vanhan ristinpuun','white', False)
#
#
#sivu = Slide(messu,True)
#sivu.InsertSectionTitle('Johdanto')
#sivu.UpdateToc('Johdanto','Alkujuonto ja seurakuntalaisen sana')
#
#
#sivu = Slide(messu,True)
#sivu.InsertSectionTitle('Sana')
#sivu.UpdateToc('Sana','päivän laulu')
#
#sivu = Slide(messu,True)
#sivu.InsertSectionTitle('Sana')
#sivu.UpdateToc('Sana','evankeliumi')
#
#sivu = Slide(messu,True)
#sivu.InsertSectionTitle('Sana')
#sivu.UpdateToc('Sana','saarna')
#
#sivu = Slide(messu,True)
#sivu.InsertSectionTitle('Sana')
#sivu.UpdateToc('Sana','synnintunnustus ja synninpäästö')
