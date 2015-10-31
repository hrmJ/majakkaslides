import uno
import sys
from com.sun.star.awt import Size
from com.sun.star.awt import Point
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.awt.FontWeight import BOLD
from collections import OrderedDict
from sqlalchemy import create_engine, func
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base
from songstodb import Song, Verse
import re
import glob
from difflib import SequenceMatcher
import menus
import pysword
import pathlib
from PIL import Image


engine = create_engine('sqlite:////home/juho/Dropbox/srk/laulut.db', echo=False)
Base = declarative_base()
scriptpath = '/home/juho/projects/majakkaslides/'

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

    def DeletePageAtEnd(self):
        """Create new page at the end"""
        self.document.DrawPages.remove(self.curpage)

class Slide:

    def __init__(self, pres, is_wide=False,indent=None,customwidth=None,vertindent=None):
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
        self.totalwidth = self.body.Size.Width
        if indent:
            self.location.X = indent
        if vertindent:
            self.location.Y = vertindent
        if is_wide:
            self.size.Width = self.body.Size.Width-1000
        else:
            self.size.Width = self.body.Size.Width-8000

        if customwidth:
            self.size.Width = self.body.Size.Width-customwidth
        

        self.size.Height = self.body.Size.Height
        self.body.setPosition(self.location)
        self.body.setSize(self.size)

    def InsertHeader(self, headertext, color, italics = False, smallfont = False, customsize=None, bold=False,center=None):
        color = CodeColor(color)
        #self.cursor.ParaAdjust = 3	
        self.cursor.CharColor = color
        self.cursor.CharHeight=26
        if smallfont:
            self.cursor.CharHeight=18
        if customsize:
            self.cursor.CharHeight=customsize
        if italics:
            self.cursor.CharPosture = ITALIC
        if bold:
            self.cursor.CharWeight = BOLD
        if center:
            self.cursor.ParaAdjust = 3
        else:
            self.cursor.CharPosture = NONE
        self.text.insertString(self.cursor, headertext + '\n',False) 	


    def InsertHeaderImage(self,pres,path):
        path = scriptpath + path
        self.drawpage = pres.curpage
        self.image = pres.document.createInstance("com.sun.star.drawing.GraphicObjectShape")
        imagelocation = Point()
        imagesize = Size()
        imagelocation.X = 18000
        imagelocation.Y = 700
        #get original image size
        self.image.Position = imagelocation
        self.image.GraphicURL = pathlib.Path(path).as_uri()
        im=Image.open(path)
        #convert to ooo size
        self.imagesize = im.size
        imagesize.Width = int(26.45833*im.size[0])
        imagesize.Height = int(26.45833*im.size[1])
        self.image.Size = imagesize
        self.drawpage.add(self.image)

    def InsertHeaderImage2(self,pres,path,imgsize,customx=None, customy=None):
        self.drawpage = pres.curpage
        self.image = pres.document.createInstance("com.sun.star.drawing.GraphicObjectShape")
        imagelocation = Point()
        imagesize = Size()
        self.image.Position = imagelocation
        self.image.GraphicURL = pathlib.Path(path).as_uri()
        imagelocation.X = self.totalwidth - imgsize[0]
        if customx:
            imagelocation.X = customx
        imagelocation.Y = 700
        if customy:
            imagelocation.Y = customy
        imagesize.Width = imgsize[0]
        imagesize.Height = imgsize[1]
        self.image.Position = imagelocation
        self.image.Size = imagesize
        self.drawpage.add(self.image)

class ImageSlide(Slide):

    def __init__(self, pres, is_wide=False,indent=None):
        path = '/home/juho/projects/majakkaslides/images/test2.png'
        super().__init__(pres,is_wide,indent)
        self.InsertHeader('Alkulaulu','black',customsize=40,bold=True)
        self.drawpage = pres.curpage
        self.image = pres.document.createInstance("com.sun.star.drawing.GraphicObjectShape")
        imagelocation = Point()
        imagesize = Size()
        imagelocation.X = 15000
        imagelocation.Y = 700
        #get original image size
        self.image.Position = imagelocation
        self.image.GraphicURL = pathlib.Path(path).as_uri()
        im=Image.open(path)
        #convert to ooo size
        self.imagesize = im.size
        imagesize.Width = int(26.45833*im.size[0])
        imagesize.Height = int(26.45833*im.size[1])
        self.image.Size = imagesize
        self.drawpage.add(self.image)


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

class PerheSongSlide(SongSlide):

    def __init__(self, pres, searched_filename, role, imgpath='',sngimgpath='',imgresize=1,tbwidt=None,sakimages=None,smresize=1):
        print('Adding ' + role)
        self.GetSongFromDb(searched_filename)
        if imgpath:
            #get original image size
            path = scriptpath + imgpath
            imgsize = GetImgSize(path,resize=imgresize)
            headerslide = Slide(pres, customwidth=imgsize[0])
        else:
            headerslide = Slide(pres, customwidth=6000)

        if sngimgpath:
            #get the songs img size
            spath = scriptpath + sngimgpath
            im=Image.open(sngimgpath)
            simgsize = GetImgSize(spath,imgresize)

        if role == 'simple':
            pass
        elif role:
            headerslide.InsertHeader(role.upper() + '\n',color='black',customsize=40,bold=True)
            headerslide.InsertHeader(self.title.upper(),color='black',customsize=40,bold=True)
            headerslide.InsertHeaderImage2(pres,path,imgsize)
            #headerslide.InsertHeader(self.title,'black',customsize=40)
        else:
            #if no role specified, insert just the yellow text
            headerslide.InsertHeader(self.title,'black',True)

        for idx, sak in enumerate(self.sakeistot):
            newslide = Slide(pres, False,customwidth=simgsize[0])
            newslide.InsertHeader(sak[0].upper(), color='black', customsize=35)
            if sakimages:
                spath = scriptpath + sakimages[idx]
                simgsize = GetImgSize(spath,smresize)
                newslide.InsertHeaderImage2(pres, spath, simgsize)
            else:
                newslide.InsertHeaderImage2(pres, spath, simgsize)

class PraiseSongSlide(SongSlide):

    def __init__(self,pres, searched_filename, role, moreinfo=''):
        ylinfo1 = 'Ylistys- ja rukouslaulujen aikana voit kirjoittaa omia  rukousaiheitasi ja hiljentyä sivualttarin luona.'
        ylinfo2 = 'Rukouspalvelu hiljaisessa huoneessa.'
        InfoSlide(pres,ylinfo1,ylinfo2)
        super().__init__(pres, searched_filename, role, moreinfo='')

def BibleSlide(pres, title='', text='',textcolor='white',headercolor='yellow', bigfont=False):
    headerslide = Slide(pres, False)
    if title:
        headerslide.InsertHeader(title,headercolor,True)
    headerslide.InsertHeader(text.address, textcolor)
    for verse in text.verselist:
        newslide = Slide(pres, False)
        if bigfont:
            newslide.InsertHeader(verse.upper(), textcolor)
        else:
            newslide.InsertHeader(verse, textcolor, customsize=35)

def CodeColor(color):
    if color == 'white':
        return 16777215
    elif color == 'yellow':
        return 13421568
    elif color == 'red':
        return 8388608
    elif color == 'black':
        return -1

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

def LuoPerheMessu(songs,evankeliumi=None):
    messu = Presentation()

    PerheOtsikko(messu,'TERVETULOA\nPERHEMESSUUN!',scriptpath + 'images/pomppu.png')
    PerheSongSlide(messu, songs['alkulaulu'], 'Alkulaulu', imgpath='images/laulu_iso.png',sngimgpath='images/laulu_pieni.png')
    KokoonnummeSlide(messu)
    PerheOtsikko(messu,'\n\nRISTINMERKKI',scriptpath + 'images/ristinmerkki.png')

    PerheOtsikko(messu,'\n\nLUETAAN RAAMATTUA',scriptpath + 'images/raamattu.png')
    BibleSlide(messu,text=evankeliumi,textcolor='black',headercolor='black',bigfont=True)

    PerheSongSlide(messu, songs['paivanlaulu'], 'PÄIVÄN LAULU', imgpath='images/laulu_iso.png',sngimgpath='images/laulu_pieni.png')

    PerheOtsikko(messu,'SAARNA:\nKUUNNELLAAN JA KATSELLAAN',scriptpath + 'images/kuunnellaanjakatsellaan.png')

    SyntiSlide(messu)
    PerheSongSlide(messu, songs['synnintunnustuslaulu'], 'SYNNIN-\nTUNNUSTUS-\nLAULU', imgpath='images/synnint.png',sngimgpath='images/laulu_pieni.png')
    PerheOtsikko(messu, 'JUMALA\nANTAA\nANTEEKSI',scriptpath + 'images/synnint.png',isbold=False,center=True,indent=4800,vertindent=2000,customy=1500)

    PerheOtsikko(messu, 'USKON-\nTUNNUSTUS-\nLAULU',scriptpath + 'images/uskontunnustus.png',isbold=False,center=True,indent=4800,vertindent=2000,customy=500)
    PerheSongSlide(messu, songs['uskontunnustuslaulu'],role='simple', sngimgpath='images/laulu_pieni.png')

    PerheOtsikko(messu, 'RUKOILLAAN',scriptpath + 'images/rukous.png',isbold=False,center=True,indent=3800,vertindent=6000,customy=1500)
    PerheSongSlide(messu, songs['rukouslaulu'], role='simple', sngimgpath='images/laulu_pieni.png')

    KolehtiSlide(messu)

    PerheOtsikko(messu, 'EHTOOLLIS-\nRUKOUS',scriptpath + 'images/ehtoollisrukous.png',isbold=False,center=False,indent=1400,vertindent=3000,customy=3500,customsize=0.6)
    PerheOtsikko(messu, 'ISÄ MEIDÄN',scriptpath + 'images/rukous.png',isbold=False,center=False,indent=1400,vertindent=3000,customy=3500,customsize=0.8)
    PerheSongSlide(messu, 'isä meidän perhemessu', 'simple', imgpath='images/rukous.png',sngimgpath='images/rukous.png',imgresize=0.7)
    PerheOtsikko(messu, 'EHTOOLLINEN\ntai\nSIUNAAMINEN',scriptpath + 'images/ehtoollinen.png',isbold=False,center=True,indent=1300,vertindent=3000,customy=1500,customsize=0.9)

    PerheSongSlide(messu, 'jumalan karitsa Riemumessusta_perhemessun_versio', ' ',imgpath='images/laulu_iso.png', sngimgpath='images/laulu_pieni.png')

    for elaulu in songs['ehtoollislaulut']:
        PerheSongSlide(messu, elaulu, 'EHTOOLLIS-\nLAULUJA:', imgpath='images/laulu_iso.png',sngimgpath='images/laulu_pieni.png')


    sakimages = ['images/taputus.png','images/jalat.png','images/kumarrus.png','images/pomppu.png','images/taputus.png']
    PerheSongSlide(messu, songs['kiitoslaulu'], 'KIITOSLAULU', imgpath='images/laulu_iso.png',sngimgpath='images/laulu_pieni.png',sakimages=sakimages,smresize=0.6)

    PerheOtsikko(messu, 'SIUNAAMINEN',scriptpath + 'images/siunaus.png',isbold=False,center=True,indent=1800,vertindent=6000,customy=1500)


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

def ExtractStructure_perhe(mailfile):
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

    match = re.search(r'Ehtoollislaulut.*\n ?--+\n(([a-öA-Ö].*\n)+)',structure)
    ehtoollislaulut = match.group(1)
    songs['ehtoollislaulut'] = ehtoollislaulut.splitlines()


    match = re.search(r'Kiitoslaulu: ?(.*)',structure)
    songs['kiitoslaulu'] = match.group(1).strip()

    match = re.search(r'Rukouslaulu: ?(.*)',structure)
    songs['rukouslaulu'] = match.group(1).strip()

    match = re.search(r'uskontunnustuslaulu: ?(.*)',structure)
    songs['uskontunnustuslaulu'] = match.group(1).strip()

    match = re.search(r'Jumalan karitsa: ?(.*)',structure)
    songs['jumalankaritsa'] = match.group(1).strip()

    match = re.search(r'Synnintunnustuslaulu: ?(.*)',structure)
    songs['synnintunnustuslaulu'] = match.group(1).strip()

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
        LuoPerheMessu(songs, evankeliumi)

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

def GetImgSize(path,resize=None):
    im=Image.open(path)
    if resize:
        return [int(26.45833*(im.size[0]*resize)),int(26.45833*(im.size[1]*resize))]
    else:
        return [int(26.45833*im.size[0]),int(26.45833*im.size[1])]


def KokoonnummeSlide(pres):
    metatext = 'ME KOKOONNUMME ISÄN JA POJAN JA PYHÄN HENGEN NIMESSÄ.'
    imgpaths = ['images/isa.png','images/poika.png','images/pyhahenki.png']
    headerslide = Slide(pres,customwidth=1)
    headerslide.InsertHeader(metatext,color='black',customsize=40,bold=True,center=True)
    customx = 3000
    for path in imgpaths:
        path = scriptpath + path
        size = GetImgSize(path)
        headerslide.InsertHeaderImage2(pres,path,GetImgSize(path),customx,customy=6500)
        customx += size[0] + 1000

def SyntiSlide(pres):
    headerslide = Slide(pres,customwidth=1)
    customyvals = [1500,1000,1500,7000,6500]
    previmg=500
    for idx, path in enumerate(glob.glob(scriptpath + 'images/synti*png')):
        size = GetImgSize(path,resize=0.7)
        customy = customyvals[idx]
        headerslide.InsertHeaderImage2(pres,path,size,previmg,customy)
        if idx==2:
            previmg = 4500
        else:
            previmg += size[0] 

def KolehtiSlide(pres):
    metatext = 'KOLEHTIKOHDE:'
    headerslide = Slide(pres,customwidth=1,indent=1000)
    headerslide.InsertHeader(metatext,color='black',customsize=40,bold=True,center=True)
    previmg=1500
    for idx, path in enumerate(['images/kimbilio.png','images/kolehtikippo.png']):
        path = scriptpath + path
        size = GetImgSize(path,resize=0.6)
        headerslide.InsertHeaderImage2(pres,path,size,previmg,customy=4000)
        previmg += size[0] 

def PerheOtsikko(pres,metatext,path,isbold=True,center=False,indent=None,vertindent=700,customy=None,customsize=1):
    size = GetImgSize(path)
    #headerslide = Slide(pres, customwidth=size[0],indent=indent,vertindent=vertindent)
    headerslide = Slide(pres, indent=indent,vertindent=vertindent)
    headerslide.InsertHeader(metatext,color='black',customsize=40,bold=isbold,center=center)
    headerslide.InsertHeaderImage2(pres,path,GetImgSize(path,customsize),customy=customy)

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


#if __name__ == "__main__":
try:
    if sys.argv[1] == 'perhe':
        messu = Presentation()
        ExtractStructure_perhe(sys.argv[2])
    else:
        mailfile = sys.argv[1]
        ExtractStructure(mailfile)
except IndexError:
    print('Usage: {} <mailfile name>'.format(sys.argv[0]))

