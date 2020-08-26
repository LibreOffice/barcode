
def rgb( r, g, b ):
	return 65536 * int( r * 255 ) + 256 * int( g * 255 ) + int( b * 255 )

def RGB( R, G, B ):
	return 65536 * R + 256 * G + B

def setpos( shape, x, y, w = None, h = None ):
	if w is not None:
		size = shape.Size
		size.Width, size.Height = w, h
		shape.Size = size
	position = shape.Position
	position.X, position.Y = x, y
	shape.Position = position

def createShape( model, page, shapetype, color = None ):
	shape = model.createInstance( 'com.sun.star.drawing.%sShape'%shapetype )
	page.add( shape )
	if color is not None:
		shape.FillStyle = 'SOLID'
		shape.FillColor = color
	elif hasattr( shape, 'FillStyle' ):
		shape.FillStyle = 'NONE'
	return shape

def createPolygon( model, page, coordss, color = None, type = 'PolyPolygon' ):
	shape = createShape( model, page, type, color )

	from com.sun.star.awt import Point
	lines = []
	for coords in coordss:
		line = []
		for x, y in coords:
			p = Point()
			p.X = int( x )
			p.Y = -int( y )
			line.append( p )
		lines.append( tuple( line ) )
	shape.PolyPolygon = tuple( lines )
	return shape

def embed( doc, imagefilename ):
	com_sun_star_text_TextContentAnchorType_AS_CHARACTER = 1
	import random
	import unohelper
	internalname = 'embeddedimage_%d'%random.randint( 100000, 999999 )
	try:
		# Writer
		graphic = doc.createInstance( 'com.sun.star.text.TextGraphicObject' )
	except:
		# otherwise
		graphic = doc.createInstance( 'com.sun.star.drawing.GraphicObjectShape' )
	table = doc.createInstance( 'com.sun.star.drawing.BitmapTable' )
	table.insertByName( internalname, unohelper.systemPathToFileUrl( imagefilename ) )
	graphic.GraphicURL = table.getByName( internalname )
	try:
		# Writer
		graphic.AnchorType = com_sun_star_text_TextContentAnchorType_AS_CHARACTER
	except:
		# otherwise
		pass
	return graphic
