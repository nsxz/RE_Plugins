package 
{
	import flash.display.*;
	import flash.text.*;
	import flash.utils.*;
	import flash.net.*;
	
	public class Util
	{
		private static var myTextBox:TextField = new TextField(); 
		private static var init:Boolean = false;
		private static var cnt:int = 0;
		public static var ignoreX:int = 0;
		
		public function Util() {
			
		}
		
		public static function doInit(s:Sprite):void {
			init = true;
			myTextBox.width = 600; 
            myTextBox.height = 600; 
            myTextBox.multiline = true; 
            myTextBox.wordWrap = true; 
            myTextBox.border = true; 
			
			var f:TextFormat = new TextFormat(); 
			f.font = "Courier"; 
			f.size = 48; 
			myTextBox.defaultTextFormat = f; 
			
			s.addChild(myTextBox); 
		}
		
		public static function DumpMessage(message:String):void
		{
			cnt++;
			if (ignoreX != 0) {
					if (cnt <= ignoreX) return;
			}
			
			if (!init)
			{
				unescape(message);  //assume flash monitor is hooking this...
			} 
			else{
				myTextBox.text += cnt.toString() + ":\n-------------------------------------\n" + message + "\n\n"; //for visual display
			}
		}

		public static function BA2Hex(buffer:ByteArray):void
		{
			var lines:String = "";
			var l:int = buffer.length;
			var origPosition:uint = buffer.position;
			buffer.position = 0;
			
			for (var j:int = 0; j < buffer.length; j++)
			{
				var value:int = buffer.readUnsignedByte();
				lines += fillUp(value.toString(16).toUpperCase(), 2, "0");
			}
			
			buffer.position = origPosition;
			DumpMessage(lines);
		}
		public static function DumpByteArray(buffer:ByteArray):void
		{
			var lines:String = fillUp("Offset", 8, " ") + "  00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F\n";
			var offset:int = 0;
			var l:int = buffer.length;
			var origPosition:uint = buffer.position;
			buffer.position = 0;
			
			for (var i:int = 0; i < l; i += 16)
			{
				lines += fillUp(offset.toString(16).toUpperCase(), 8, "0") + "  ";
				var line_max:int = Math.min(16, buffer.length - buffer.position);
				var ascii_line:String = "";

				for (var j:int = 0; j < 16; ++j)
				{
					if (j < line_max)							
					{
						var value:int = buffer.readUnsignedByte();
						ascii_line += value >= 32 ? String.fromCharCode(value) : ".";
						lines += fillUp(value.toString(16).toUpperCase(), 2, "0") + " ";
						offset++;
					}
					else
					{
						lines += "   ";
					}
				}
				lines += " " + ascii_line + "\n";
			}

			buffer.position = origPosition;
			DumpMessage(lines);
		}
				
		private static function fillUp(value:String, count:int, fillWith:String):String
		{
			var l:int = count - value.length;
			var ret:String = "";
			while (--l > -1)
				ret += fillWith;
			return ret + value;
		}
	}
}
