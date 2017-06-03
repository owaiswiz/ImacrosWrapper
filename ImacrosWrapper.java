package imacroswrapper;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Variant;
/**
 *
 * @author owaiswiz
 */
class ImacrosWrapper {
    Variant args[] = new Variant[2];
    ActiveXComponent iim;
    
    public ImacrosWrapper() 
    {
        iim = new ActiveXComponent("imacros"); 
    }
    public Variant iimOpen()
    {
        return iim.invoke("iimOpen");
    }
    public Variant iimOpen(String cmdLine)
    {
        return iim.invoke("iimOpen",cmdLine);   
    }
    
    public Variant iimGetInterfaceVersion()
    {
        return iim.invoke("iimGetInterfaceVersion");
    }
    
    public Variant iimPlay(String macro,int timeout)
    {
        return iim.invoke("iimPlay",macro,timeout);
    }
    public Variant iimPlay(String macro)
    {
        return iimPlay(macro,600);
    }
    
    public Variant iimPlayCode(String macroCode,int timeout)
    {
        return iimPlay("CODE:" + macroCode,timeout);
    }
    public Variant iimPlayCode(String macroCode)
    {
        return iimPlayCode(macroCode,600);
    }
    
    public Variant iimDisplay(String displayString)
    {
        return iimDisplay(displayString,10);
    }
    public Variant iimDisplay(String displayString,int timeout)
    {
        return iim.invoke("iimDisplay",displayString,timeout);
    }
	
    public Variant iimSet(String varName,Object varVal)
    {
        initVariantArgs(varName, varVal);
        return iim.invoke("iimSet",args);
    }
    public Variant iimSetInternal(String varName,Variant varVal)
    {
        initVariantArgs(varName, varVal);
        return iim.invoke("iimSetInternal",args);
    }
    
    public Variant iimGetExtract()
    {
        return iimGetExtract(0);
    }
    public Variant iimGetExtract(int index)
    {
        return iim.invoke("iimGetExtract",index);
    }
    
    public Variant iimGetErrorText()
    {
        return iim.invoke("iimGetErrorText");
    }
    
    public Variant iimClose()
    {
        return iim.invoke("iimClose");
    }
    
    public Variant invoke(String cmd)
    {
        return iim.invoke(cmd);
    }
    public Variant invoke(String cmd,String param)
    {
        return iim.invoke(cmd,param);
    }
    public Variant invoke(String cmd,String param1,int param2)
    {
        return iim.invoke(cmd,param1,param2);
    }
    public Variant invoke(String cmd,String param1,String param2)
    {
        initVariantArgs(param1, param2);
        return iim.invoke(cmd,args);
    }
	public Variant invoke(String cmd,String param1,Boolean param2)
	{
		initVariantArgs(param1,param2);
		return iim.invoke(cmd,args);
	}
	public Variant invoke(String cmd, Variant args[])
	{
		return iim.invoke(cmd,args);
	}
    
    private void initVariantArgs(Object varName, Object varVal)
    {
        args[0] = new Variant(varName);
        args[1] = new Variant(varVal);
    } 
}