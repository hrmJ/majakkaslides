import glob
import re
import datetime

from sqlalchemy import create_engine, ForeignKey
from sqlalchemy import Column, Date, Integer, String, Float
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship, backref

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

import datetime

engine = create_engine('sqlite:////home/juho/Dropbox/srk/laulut.db', echo=False)
Base = declarative_base()


class Song(Base):
    """Dishes consist of ingredients. """
    __tablename__ = "songs"
 
    id = Column(Integer, primary_key=True)
    title = Column(String)  
    sav = Column(String)  
    san = Column(String)  
    added = Column(Date)

    def __init__(self, title='', sav='', san='', suomsan=''):
        self.title = title
        self.sav = sav
        self.san = san
        self.suomsan = suomsan
        self.added = datetime.datetime.today()


class Verse(Base):
    """Ingredients are linked to dishes by the linkid column"""
    __tablename__ = "verses"
 
    id = Column(Integer, primary_key=True)
    content = Column(String)  
    # If you want to clssify ingerdients some way
    versetype = Column(String)
    song_id = Column(Integer, ForeignKey("songs.id"))
    song = relationship("Song", backref=backref("verses", order_by=id))
 
    def __init__(self, content, versetype = 'unspecified' ):
        """initialize so that column names don't have to be specified"""
        self.content = content
        self.versetype = versetype


#=====================================================================

if __name__ == "__main__":
    Base.metadata.create_all(engine)

    songpath = '/home/juho/Dropbox/laulut/*.txt'

    Session = sessionmaker(bind=engine)
    session = Session()

    for songfile in glob.glob(songpath):
        upcount = 0
        if not 'laulut.txt' in songfile and not '_laulujen lista.txt' in songfile:
            with open(songfile,'r') as f:
                laulu_raw = f.read()
                #Split if 2 or more consequtive unix line breaks
                verses = list()
                verses = re.split('\n{2,}',laulu_raw)
                thistitle = verses[0]
                res = None
                res = session.query(Song.title).filter(Song.title==thistitle).first()
                if not res and 'euvonen' not in thistitle and 'suom. san' not in thistitle:
                    #If no previous song by this name
                    sav = ''
                    san = ''
                    suomsan = ''
                    if "säv" in thistitle and "san" in thistitle:
                        thistitle = input('Varmista laulun nimi: {}\n>'.format(thistitle))
                        sav = input('Sävel:')
                        san = input('Sanat:')
                        suomsan = input('Suom.sanat:')
                    laulu = Song(thistitle,sav,san,suomsan)
                    laulu.verses = []
                    for verse in verses[1:]:
                        laulu.verses.append(Verse(content=verse))
                    session.add(laulu)
                    upcount += 1
        session.commit()

    print('Done. Updated {} new songs.'.format(upcount))
