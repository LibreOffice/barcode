
if 'set' not in globals():
    from sets import Set as set
if 'sorted' not in globals():
    def sorted( l ):
        l = list( l )
        l.sort()
        return l

MAX_CACHE_SIZE = 10 * 1024 * 1024

cachefile = './cache'
def setcachefile( filename ):
    global cachefile
    cachefile = filename

class WebCache( object ):
    def __init__( self, filename = None ):
        import os.path
        if filename is None:
            filename = cachefile
        if os.path.exists( filename ):
            try:
                import os
                if os.stat( filename ).st_size > MAX_CACHE_SIZE:
                    raise Exception( 'cache grew too big' )
                self.data = eval( '{' + file( filename ).read() + '}' )
            except:
                self.data = {}
                self.file = file( filename, 'w' )
            else:
                self.file = file( filename, 'a' )
        else:
            self.data = {}
            self.file = file( filename, 'w' )
    def getpage( self, url ):
        url = url.encode( 'ascii' )
        if url not in self.data:
            html = loadpage( url )
            self.data[url] = html
            self.file.write( "'%s': '%s',\n"%(url.encode( 'string_escape' ), html.encode( 'string_escape' )) )
        return self.data[url]

def loadpage( url ):
    import urllib.request, urllib.error, urllib.parse
    opener = urllib.request.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0 (OpenOffice.org extension "Barcode" created with EuroOffice Extension Creator)')]
    html = opener.open( url ).read()
    return html

cache = None
def getpage( url ):
    global cache
    if cache is None:
        cache = WebCache()
    return cache.getpage( url )

def flatten( x ):
    if hasattr( x, 'contents' ):
        return ' '.join( [flatten( y ) for y in x.contents] )
    else:
        return str( x ).strip()


