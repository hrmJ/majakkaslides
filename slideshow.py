import uno
import sys
from com.sun.star.awt import Size
from com.sun.star.awt import Point
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt.FontSlant import ITALIC
from collections import OrderedDict
from sqlalchemy import create_engine, func
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from songstodb import Song, Verse
import re
from difflib import SequenceMatcher
import menus
import pysword

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

    def InsertHeader(self, headertext, color, italics = False, smallfont = False):
        color = CodeColor(color)
        #self.cursor.ParaAdjust = 3	
        self.cursor.CharColor = color
        self.cursor.CharHeight=26
        if smallfont:
            self.cursor.CharHeight=18
        if italics:
            self.cursor.CharPosture = ITALIC
        else:
            self.cursor.CharPosture = NONE
        self.text.insertString(self.cursor, headertext + '\n',False) 	

class InfoSlide(Slide):
    """Slides that give some general info"""

    def __init__(self, presentation, infotext1, infotext2):
        super().__init__(presentation, False)
        self.cursor.ParaLeftMargin = 0
        self.InsertHeader(infotext1,'yellow', True)
        self.InsertHeader(infotext2,'white', False)

class Metaslide(Slide):
    """Slides that give information about the current stage of the service"""

    def __init__(self, presentation, currentsect, currenttitle):
        super().__init__(presentation, True)
        self.UpdateToc(currentsect, currenttitle)


    def InsertSectionTitle(self, current):
        sections = ['Johdanto','Sana','Ylistys ja rukous','Ehtoollinen','Siunaus ja lähettäminen']
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
        self.InsertSectionTitle(currentsect)
        sections = OrderedDict()
        sections['Johdanto'] = ['Alkulaulu', 'Alkusanat ja seurakuntalaisen sana']
        sections['Sana'] = ['päivän laulu','evankeliumi','saarna','synnintunnustus','uskontunnustus']
        sections['Ylistys ja rukous'] = ['Ylistys- ja rukouslauluja','Esirukous']
        sections['Ehtoollinen'] = ['Pyhä','Ehtoollisrukous', 'Isä meidän', 'Jumalan karitsa', 'Ehtoollisen jako']
        sections['Siunaus ja lähettäminen'] = ['Herran siunaus','Loppusanat','Loppulaulu']

        self.text.insertString(self.cursor, '\n'*2, False) 	
        self.cursor.CharHeight=20

        for title in sections[currentsect]:
            if title == currenttitle:
                title = ' --> ' + title
            else:
                title = '       ' + title
            self.text.insertString(self.cursor, title + '\n',False) 	
            self.cursor.CharColor = CodeColor('white')



class SongSlide():
    """Song header + actual song"""
    session = None

    def __init__(self, pres, searched_filename, role, moreinfo=''):
        print('Adding ' + role)
        self.GetSongFromDb(searched_filename)
        if moreinfo:
            headerslide = Slide(pres, False)
        else:
            headerslide = Slide(pres, True)
        if role:
            headerslide.InsertHeader(role,'yellow',True)
            headerslide.InsertHeader(self.title,'white')
        else:
            #if no role specified, insert just the yellow text
            headerslide.InsertHeader(self.title,'yellow',True)
        if moreinfo:
            #If additional information to be followed:
            headerslide.InsertHeader(moreinfo,'white',True, True)
        for sak in self.sakeistot:
            newslide = Slide(pres, True)
            newslide.InsertHeader(sak[0],'white')


    def GetSongFromDb(self, searched_filename):
        """Search the database for songs with this name"""
        if not SongSlide.session:
            Session = sessionmaker(bind=engine)
            SongSlide.session = Session()
        try:
            session = SongSlide.session
            laulu = session.query(Song)
            searched_filename = searched_filename.lower()
            subquery = session.query(Song.id).filter(func.lower(Song.filename) == searched_filename).subquery()
            query = session.query(Verse.content).filter(Verse.song_id.in_(subquery))
            self.sakeistot = query.all()
            title = session.query(Song.title).filter(func.lower(Song.filename) == searched_filename).first()
            self.title = title[0]
        except TypeError:
            input('ei löydy laulua ' + searched_filename)
            self.title = searched_filename
            self.sakeistot = [('laulupuuttuu..')]

class PraiseSongSlide(SongSlide):

    def __init__(self,pres, searched_filename, role, moreinfo=''):
        ylinfo1 = 'Ylistys- ja rukouslaulujen aikana voit kirjoittaa omia  rukousaiheitasi ja hiljentyä sivualttarin luona.'
        ylinfo2 = 'Rukouspalvelu hiljaisessa huoneessa.'
        InfoSlide(pres,ylinfo1,ylinfo2)
        super().__init__(pres, searched_filename, role, moreinfo='')

def BibleSlide(pres, title, text):
    headerslide = Slide(pres, False)
    headerslide.InsertHeader(title,'yellow',True)
    headerslide.InsertHeader(text.address, 'white')
    for verse in text.verselist:
        newslide = Slide(pres, False)
        newslide.InsertHeader(verse, 'white')

def CodeColor(color):
    if color == 'white':
        return 16777215
    elif color == 'yellow':
        return 13421568

def LuoOtsikkodia(ThisComponent, yellowtext, whitetext):
    #Create a new slide at the end
    ThisComponent.DrawPages.insertNewByIndex(ThisComponent.DrawPages.Count)


def LuoMessu(songs,evankeliumi=None):
    plinfo1 = 'Päivän laulun aikana 3-6-vuotiaat lapset voivat siirtyä pyhikseen ja yli 6-vuotiaat klubiin.' 
    plinfo2 = 'Seuraa vetäjiä - tunnistat heidät lyhdyistä!'

    kolinfo1 = 'Voit tulla ehtoolliselle jo Jumalan karitsa -hymnin aikana' 
    kolinfo2 = 'Halutessasi voit jättää kolehdin ehtoolliselle tullessasi oikealla olevaan koriin.'

    messu = Presentation()
    #Johdanto
    Metaslide(messu,'Johdanto','Alkulaulu')
    SongSlide(messu, songs['alkulaulu'], 'Alkulaulu')
    Metaslide(messu,'Johdanto','Alkusanat ja seurakuntalaisen sana')
    ##Sana
    Metaslide(messu,'Sana','päivän laulu')
    InfoSlide(messu, plinfo1,plinfo2)
    SongSlide(messu, songs['paivanlaulu'], 'Päivän laulu')
    Metaslide(messu,'Sana','evankeliumi')
    if evankeliumi:
        BibleSlide(messu, 'Evankeliumi / raamatunkohta', evankeliumi)
    Metaslide(messu,'Sana','saarna')
    Metaslide(messu,'Sana','synnintunnustus')
    Metaslide(messu,'Sana','uskontunnustus')
    SongSlide(messu, 'uskontunnustus', '')
    ##Rukous
    Metaslide(messu,'Ylistys ja rukous','Ylistys- ja rukouslauluja')
    for ylistyslaulu in songs['ylistyslaulut']:
        PraiseSongSlide(messu, ylistyslaulu, 'Ylistys- ja rukouslauluja')
    Metaslide(messu,'Ylistys ja rukous','Esirukous')
    #Ehtoollinen
    Metaslide(messu,'Ehtoollinen','Pyhä')
    SongSlide(messu, songs['pyha'], 'Pyhä')
    Metaslide(messu,'Ehtoollinen','Ehtoollisrukous')
    InfoSlide(messu, '', '')
    SongSlide(messu, 'isä meidän', '')
    Metaslide(messu,'Ehtoollinen','Jumalan karitsa')
    InfoSlide(messu, kolinfo1, kolinfo2)
    SongSlide(messu, songs['jumalankaritsa'], '')
    for ehtoollislaulu in songs['ehtoollislaulut']:
        InfoSlide(messu, '', '')
        SongSlide(messu, ehtoollislaulu, 'Ehtoollislauluja')
    #Lähettäminen
    Metaslide(messu,'Siunaus ja lähettäminen','Herran siunaus')
    Metaslide(messu,'Siunaus ja lähettäminen','Loppusanat')
    SongSlide(messu, songs['loppulaulu'], 'Loppulaulu')
    print('Done. Muista otsikko, kolehtidia ja...')

def ExtractStructure(mailfile):
    with open(mailfile,'r') as f:
        structure = f.read()

    songs = dict()

    match = re.search(r'Alkulaulu: ?(.*)',structure)
    songs['alkulaulu'] = match.group(1).strip()

    match = re.search(r'Päivän laulu: ?(.*)',structure)
    songs['paivanlaulu'] = match.group(1).strip()

    match = re.search(r'evankeliumi: ?(.*)',structure.lower())
    if match:
        address = match.group(1).strip()
        match = re.search(r'(\d?\w+) (\d+):([0-9,-]+)',structure.lower())
        book = match.group(1)
        chapter = match.group(2)
        verse = match.group(3)
        evankeliumi = BibleText(book,chapter,verse,address)

    match = re.search(r'Ylistyslaulut.*\n ?--+\n(([a-öA-Ö].*\n)+)',structure)
    ylistyslaulut = match.group(1)
    songs['ylistyslaulut'] = ylistyslaulut.splitlines()

    match = re.search(r'Ehtoollislaulut.*\n ?--+\n(([a-öA-Ö].*\n)+)',structure)
    ehtoollislaulut = match.group(1)
    songs['ehtoollislaulut'] = ehtoollislaulut.splitlines()

    match = re.search(r'Pyhä-hymni: ?(.*)',structure)
    songs['pyha'] = match.group(1).strip()

    match = re.search(r'Jumalan karitsa: ?(.*)',structure)
    songs['jumalankaritsa'] = match.group(1).strip()

    match = re.search(r'Loppulaulu: ?(.*)',structure)
    songs['loppulaulu'] = match.group(1).strip()

    for songrole, songname in songs.items():
        if songrole not in ('ylistyslaulut','ehtoollislaulut'):
            songs[songrole] = CheckAvailability(songname)
        else:
            #ylistslaulut, ehtoollislaulut are lists that contain many song names
            newsongnames = list()
            for thissongname in songname:
                newsongnames.append(CheckAvailability(thissongname))
            songs[songrole] = newsongnames

    cont = menus.multimenu({'y':'yes','n':'no'}, 'All songs found in the database. Create slides?')
    if cont.answer == 'y':
        LuoMessu(songs, evankeliumi)

def CheckAvailability(songname):
    """Check if this song is in the db and try to guess if not"""
    Session = sessionmaker(bind=engine)
    session = Session()
    songname = songname.lower()

    if not session.query(Song.filename).filter(func.lower(Song.filename) == songname).first():
        allnames = session.query(Song.filename).all()
        suggestions = dict()
        for name in allnames:
            simratio = SequenceMatcher(None, songname, name[0]).ratio()
            suggestions[simratio] = name[0]
        ratios = sorted(suggestions.keys())
        ratios = ratios[-10:]
        ratios = sorted(ratios[-10:],reverse=True)
        suglist = dict()
        for idx, ratio in enumerate(ratios):
            suglist[str(idx)] = suggestions[ratio]
        suglist['n'] = 'ei mikään näistä'
        fuzzymenu = menus.multimenu(suglist, promptnow = 'Vastaako jokin näistä haettavaa laulua ({})?'.format(songname))
        if fuzzymenu.answer != 'n':
            return suglist[fuzzymenu.answer]
        else:
            sys.exit('Song "{}" not found. Exiting.'.format(songname))
        print('False')

    return songname

class BibleText:

    def __init__(self, book, chapter, verse, finaddress=''):
        self.address = input('kirjan nimi *' + book + '* suomeksi?\n>') + chapter + ': ' + verse
        self.GetBibleText(book,chapter,verse)

    def GetBibleText(self, book, chapter, verse):
        """Get bible text from a sword module using the pysword library (https://github.com/kcarnold/pysword)"""
        try:
            module = pysword.ZModule('finpr92')
        except:
            sys.exit('Please set the path of the sword module in pysword.py')
        text = ''
        if '-' in verse:
            verses = verse.split('-')
            start = int(verses[0])
            end = int(verses[1])
            verses = range(start,end+1)
            pair = ''
            verselist = list()
            for verse in verses:
                if not pair:
                    pair = module.text_for_ref(book, chapter, verse).decode("utf-8") + '\n'
                else:
                    pair += module.text_for_ref(book, chapter, verse).decode("utf-8") + '\n'
                    verselist.append(pair)
                    pair = ''
            if pair:
                # jos pariton märä jakeita
                verselist.append(pair)
        else:
            verselist = [module.text_for_ref(book, chapter, verse).decode("utf-8")]
        self.verselist = verselist


if __name__ == "__main__":
    try:
        mailfile = sys.argv[1]
        ExtractStructure(mailfile)
    except IndexError:
        print('Usage: {} <mailfile name>'.format(sys.argv[0]))
