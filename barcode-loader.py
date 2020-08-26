DEBUG = False		# change to True to debug the loader

if DEBUG:
	import sys
	if sys.platform == 'win32':
		import random
		debugfile = 'c:\\debug-loader.txt' #+ str( random.randint( 100, 999 ) )
		df = file( debugfile, 'w' )
		df.close()
		del df

def debugexception():
	if not DEBUG: return
	import sys
	if sys.platform == 'win32':
		f = file( debugfile, 'a' )
	else:
		f = sys.stdout
	import traceback
	traceback.print_exc( file=f )

try:
	import os
	import uno
	import unohelper
	def typednamedvalues( type, *args, **kwargs ):
		if args:
			dict = args[0]
		else:
			dict = kwargs
		props = []
		for k, v in dict.items():
			p = uno.createUnoStruct( type )
			p.Name = k
			p.Value = v
			props.append( p )
		return tuple( props )

	def props( *args, **kwargs ):
		return typednamedvalues( 'com.sun.star.beans.PropertyValue', *args, **kwargs ) 

	def initenvironment( classname ):
		ctx = uno.getComponentContext()
		initpython( getpath( ctx, classname ) )

	def getconfig( ctx, nodepath, update = False ):
		psm = ctx.ServiceManager
		configprovider = psm.createInstance( 'com.sun.star.configuration.ConfigurationProvider' )
		configaccess = configprovider.createInstanceWithArguments( 'com.sun.star.configuration.ConfigurationAccess', props( nodepath = nodepath ) )
		return configaccess

	def getpath( ctx, classname ):
		config = getconfig( ctx, 'org.openoffice.%sSettings/ConfigNode'%classname )
		path = config.Origin
		expander = ctx.getValueByName( '/singletons/com.sun.star.util.theMacroExpander' )
		path = expander.expandMacros( path )
		path = path[len( 'vnd.sun.star.expand:' ):]
		import os
		path = unohelper.absolutize( os.getcwd(), path )
		path = unohelper.fileUrlToSystemPath( path )
		return path

	def initpython( path ):
		import sys
		if sys.hexversion >= 0x02030000 and sys.hexversion <= 0x0203ffff:
			pythonversion = '23'
		elif sys.hexversion >= 0x02050000 and sys.hexversion <= 0x0205ffff:
			pythonversion = '25'
		elif sys.hexversion >= 0x02060000 and sys.hexversion <= 0x0206ffff:
			pythonversion = '26'
		elif sys.hexversion >= 0x0206ffff:
			pythonversion = '-future'
		else:
			debug( '%s only supports Python 2.3.x, and Python 2.5 or later.'%classname )
			raise Exception( '%s only supports Python 2.3.x, and Python 2.5 or later.'%classname )
		if path not in sys.path:
			sys.path.insert( 0, path )
			import barcode_generatedsettings as gs
			if gs.DEBUG:
				# load modules from development directory in debug mode
				sys.path.pop( 0 )
				path = gs.HOME
				sys.path.insert( 0, path )
			sys.path.insert( 0, os.path.join( path, 'modules-python'+pythonversion ) )
			if sys.platform == 'win32':
				sys.path.insert( 0, os.path.join( path, 'binaries-windows-python'+pythonversion ) )
			elif sys.platform == 'darwin':
				sys.path.insert( 0, os.path.join( path, 'binaries-darwin-python'+pythonversion ) )
			else:
				sys.path.insert( 0, os.path.join( path, 'binaries-linux-python'+pythonversion ) )

	initenvironment( 'Barcode' )
	import barcode.barcode			# non-"from" import gives more relevant error message (see uno.py)
	from barcode.barcode import *
except:
	debugexception()

