package 
{
	import flash.display.Sprite;
	import flash.events.Event;
	import flash.utils.ByteArray;
	import flash.utils.*;
	
	
	
	public class Main extends Sprite 
	{

		public function Main():void 
		{			

			Util.doInit(this);
			var s:String = "my test";
			s = unescape(s);
            Util.DumpMessage(s);
			
			
		}
	
		
	}
	
}