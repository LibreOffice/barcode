import importlib
import glob
import os
import sys
import random
import traceback

import uno
import unohelper

from com.sun.star.beans import XPropertySet
from com.sun.star.task import XJob, XJobExecutor
from com.sun.star.lang import XServiceName, XInitialization, XComponent, XServiceInfo, XServiceDisplayName
from com.sun.star.uno import RuntimeException


def typednamedvalues( type, *args, **kwargs ):
    if args:
        dict = args[0]
    else:
        dict = kwargs
    props = []
    for k, v in list(dict.items()):
        p = uno.createUnoStruct( type )
        p.Name = k
        p.Value = v
        props.append( p )
    return tuple( props )

class propset( unohelper.Base, XPropertySet, XServiceInfo ):
    def __init__( self, *args, **kwargs ):
        if args:
            self.dict = args[0]
        else:
            self.dict = kwargs
        self.services = []
    # XPropertySet
    def getPropertySetInfo( self ):
        return None
    def setPropertyValue( self, name, value ):
        self.dict[name] = value
    def getPropertyValue( self, name ):
        return self.dict[name]
    def addPropertyChangeListener( self, listener ):
        pass
    def removePropertyChangeListener( self, listener ):
        pass
    def addVetoableChangeListener( self, listener ):
        pass
    def removeVetoableChangeListener( self, listener ):
        pass
    # XServiceInfo
    def getImplementationName( self ):
        return 'org.openoffice.PropertySet'
    def supportsService( self, s ):
        return s in self.services
    def getSupportedServiceNames( self ):
        return tuple( self.services )

def props( *args, **kwargs ):
    return typednamedvalues( 'com.sun.star.beans.PropertyValue', *args, **kwargs )

def anyprops( *args, **kwargs ):
    return uno.Any( '[]com.sun.star.beans.PropertyValue', props( *args, **kwargs ) )

def namedvalues( *args, **kwargs ):
    return typednamedvalues( 'com.sun.star.beans.NamedValue', *args, **kwargs )

def anynamedvalues( *args, **kwargs ):
    return uno.Any( '[]com.sun.star.beans.NamedValue', props( *args, **kwargs ) )

def enumerate( obj ):
    if hasattr( obj, 'createEnumeration' ):
        obj = obj.createEnumeration()
    if hasattr( obj, 'hasMoreElements' ):
        while obj.hasMoreElements():
            yield obj.nextElement()
    elif hasattr( obj, 'Count' ):
        for i in range( obj.Count ):
            yield obj.getByIndex( i )
    elif hasattr( obj, 'ElementNames' ):
        for n in obj.ElementNames:
            yield obj.getByName( n )

class EasyDict( dict ):
    def __getattr__( self, key ):
        return self[key]
    def __hasattr__( self, key ):
        return key in self
    def __setattr__( self, key, value ):
        self[key] = value

def unprops( props ):
    dict = EasyDict()
    for p in props:
        dict[p.Name] = p.Value
    return dict

def debug( *msgs ):
    try:
        print(msgs)
    except:
        debugexception()

def debugexception():
    traceback.print_exc( file=sys.stdout )

documenttypes = [
    'com.sun.star.frame.StartModule',
    'com.sun.star.text.TextDocument',
    'com.sun.star.sheet.SpreadsheetDocument',
    'com.sun.star.text.WebDocument',
    'com.sun.star.drawing.DrawingDocument',
    'com.sun.star.presentation.PresentationDocument',
#    'com.sun.star.chart.ChartDocument',
    'com.sun.star.formula.FormulaProperties',
        ]

runninginstance = None

class ComponentBase( unohelper.Base, XServiceName, XInitialization, XComponent, XServiceInfo, XServiceDisplayName, XJobExecutor, XJob ):
    def __init__( self, *args ):
        self.isAPIMode = False
        # store the component context for later use
        try:
            self.ctx = args[0]
            self.config = self.getconfig( 'org.libreoffice.%sSettings/ConfigNode'%self.__class__.__name__, update = True )
            self.initpath()
            self.initlanguage()
        except:
            debugexception()

    # XInitialization
    def initialize( self, args ):
        pass
    # XComponent
    def dispose( self ):
        pass
    def addEventListener( self, listener ):
        pass
    def removeEventListener( self, listener ):
        pass
    # XServiceInfo
    def getImplementationName( self ):
        try:
            return 'org.libreoffice.' + self.__class__.__name__
        except:
            debugexception()
    def supportsService( self, s ):
        return s in self.services
    def getSupportedServiceNames( self ):
        return self.services
    # XServiceDisplayName
    def getServiceDisplayName( self, locale ):
        try:
            lang = locale.Language
            if lang not in self.SUPPORTED_LANGUAGES:
                lang = self.SUPPORTED_LANGUAGES[0]
            return self.localize( 'title', language = lang )
        except:
            debugexception()

    def startup( self ):
        '''
        Runs at application startup.
        Subclasses may make use of it.
        '''
        pass

    def firstrun( self ):
        '''
        Runs at first startup after installation.
        Subclasses may make use of it.
        '''
        pass

    def coreuninstall( self ):
        try:
            self.uninstall()
        except:
            debugexception()
        self.config.FirstRun = True    # will need to run install again (in case we are reinstalled)
        self.config.commitChanges()
    def uninstall( self ):
        '''
        Runs upon uninstallation.
        Subclasses may make use of it.
        '''
        pass

    def getconfig( self, nodepath, update = False ):
        if update:
            update = 'Update'
        else:
            update = ''
        psm = self.ctx.ServiceManager
        configprovider = psm.createInstance( 'com.sun.star.configuration.ConfigurationProvider' )
        configaccess = configprovider.createInstanceWithArguments( 'com.sun.star.configuration.Configuration%sAccess'%update, props( nodepath = nodepath ) )
        return configaccess

    def initpath( self ):
        path = self.config.Origin
        expander = self.ctx.getValueByName( '/singletons/com.sun.star.util.theMacroExpander' )
        path = expander.expandMacros( path )
        path = path[len( 'vnd.sun.star.expand:' ):]
        path = unohelper.absolutize( os.getcwd(), path )
        path = unohelper.fileUrlToSystemPath( path )
        self.path = path

    def initlanguage( self ):
        config = self.getconfig( '/org.openoffice.Setup' )
        self.uilanguage = config.L10N.ooLocale.split( '-' )[0]
        if self.uilanguage not in self.SUPPORTED_LANGUAGES: self.uilanguage = self.SUPPORTED_LANGUAGES[0]

    def localize( self, string, language = None ):
        if language is None:
            language = self.uilanguage
        if not hasattr( self, 'localization' ):
            self.loadlocalization()
        if string not in self.localization: return 'unlocalized: '+string        # debug
        if language in self.localization[string]:
            return self.localization[string][language]
        elif self.SUPPORTED_LANGUAGES[0] in self.localization[string]:
            return self.localization[string][self.SUPPORTED_LANGUAGES[0]]
        else:
            return 'unlocalized for %s: %s'%(language, string)        # debug
    def loadlocalization( self ):
        self.localization = {}
        try:
            dir = 'EOEC%sDialogs'%self.__class__.__name__
            for filename in glob.glob( os.path.join( self.path, dir, 'DialogStrings_*.properties' ) ):
                sf = os.path.split( filename )[-1]
                lang = sf[sf.index( '_' )+1:sf.index( '_' )+3]
                with open(filename) as f:
                    content = f.readlines()

                for l in content:
                    l = l.split( '#' )[0].strip()
                    if len( l ) == 0: continue
                    assert '=' in l
                    key, value = l.split( '=', 1 )
                    key = key.strip()
                    value = value.strip()
                    if key not in self.localization:
                        self.localization[key] = {}
                    self.localization[key][lang] = value.replace( '\\', '' )
        except:
            debugexception()

    def trigger( self, arg ):
        try:
            getattr( self, arg )()
        except Exception:
            debugexception()

    def commandURL( self, command ):
        return 'service:org.libreoffice.%s?%s'%(self.__class__.__name__, command)

    def createdialog( self, dialogname ):
        psm = self.ctx.ServiceManager
        dlgprovider = psm.createInstance( 'com.sun.star.awt.DialogProvider' )
        dlg = dlgprovider.createDialog( 'vnd.sun.star.script:EOEC%sDialogs.%s?location=application'%(self.__class__.__name__, dialogname) )
        class Wrapper( object ):
            def __init__( self, dlg ):
                object.__setattr__( self, 'xdialog', dlg )
            def __getattr__( self, name ):
                return getattr( self.xdialog, name )
            def __setattr__( self, name, value ):
                try:
                    setattr( self.xdialog, name, value )
                except AttributeError:
                    object.__setattr__( self, name, value )
        dlg = Wrapper( dlg )
        for c in dlg.getControls():
            setattr( dlg, c.Model.Name, c )
        return dlg

    def execute( self, args ):
        try:
            args = unprops( unprops( args ).Environment )
            getattr( self, args.EventName )()
        except Exception:
            debugexception()

    def onFirstVisibleTask( self ):
        try:
            global runninginstance
            if runninginstance is None:
                runninginstance = self
            if self.config.FirstRun:
                try:
                    self.firstrun()
                except:
                    self.debugexception_and_box()
                self.config.FirstRun = False
                self.config.commitChanges()
            self.startup()
        except:
            debugexception()

    def isWriter(self):
        return self.getcomponent().supportsService("com.sun.star.text.TextDocument")

    def isDraw(self):
        return self.getcomponent().supportsService("com.sun.star.drawing.DrawingDocument")

    def isImpress(self):
        return self.getcomponent().supportsService("com.sun.star.presentation.PresentationDocument")

    def isCalc(self):
        return self.getcomponent().supportsService("com.sun.star.sheet.SpreadsheetDocument")

    def box( self, message, kind = 'infobox', buttons = 'OK', title = None ):
        # Don't show message box in API mode, throw exception instead
        if self.isAPIMode:
            raise RuntimeException(message, self.ctx)

        if kind == 'infobox' and buttons != 'OK':
            kind = 'querybox'    # infobox only supports OK
        if title is None: title = self.localize( 'title' )
        toolkit = self.ctx.ServiceManager.createInstance( 'com.sun.star.awt.Toolkit' )
        rectangle = uno.createUnoStruct( 'com.sun.star.awt.Rectangle' )
        msgbox = toolkit.createMessageBox( self.getdesktop().getCurrentFrame().getContainerWindow(),
            kind, uno.getConstantByName( 'com.sun.star.awt.MessageBoxButtons.BUTTONS_'+buttons ),
            title, message )
        return msgbox.execute()
    BOXCANCEL = 0
    BOXOK = 1
    BOXYES = 2
    BOXNO = 3
    BOXRETRY = 4

    def debugexception_and_box( self, format = None ):
        debugexception()
        try:
            if format is None:
                format = 'An unexpected error (%(kind)s) occured at line %(linenumber)s of %(filename)s.'
            tb = traceback.extract_tb( sys.exc_info()[2] )
            exc = EasyDict()
            exc.kind = sys.exc_info()[0]
            exc.filename = tb[-1][0]
            exc.linenumber = tb[-1][1]
            exc.functionname = tb[-1][2]
            exc.text = tb[-1][3]
            exc.filename = os.path.split( exc.filename )[1]
            self.box( format%exc, 'errorbox' )
        except:
            debugexception()

    def getdesktop( self ):
        psm = self.ctx.ServiceManager
        return psm.createInstanceWithContext( 'com.sun.star.frame.Desktop', self.ctx )
    def getcomponent( self ):
        d = self.getdesktop()
        c = d.getCurrentComponent()
        if c is None:
            debug( 'no currentcomponent, picking first' )
            c = d.getComponents().createEnumeration().nextElement()
        return c
    def getcontroller( self ):
        return self.getcomponent().getCurrentController()

    def getServiceName( cls ):
        try:
            return 'org.libreoffice.' + cls.__name__
        except Exception:
            debugexception()
    getServiceName = classmethod( getServiceName )
