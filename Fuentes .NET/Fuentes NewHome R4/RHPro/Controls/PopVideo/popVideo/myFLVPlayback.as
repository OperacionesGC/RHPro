package
{
    import fl.video.FLVPlayback;
    import flash.display.Sprite;
    import flash.text.TextField;
	
    public class myFLVPlayback extends Sprite {

		private var title:String;
        private var videoPath:String;/// = "http://www.helpexamples.com/flash/video/caption_video.flv";
        
        public function myFLVPlayback() {
			var params:Object = root.loaderInfo.parameters;
			
			title = params.title || null;
			videoPath = params.path || null;
			
			if (videoPath == null) {
				_title.text = "Error al recibir la ruta de video";
				return;
			}
			
			_title.htmlText = title;
			
            player.source = videoPath;
            player.skinBackgroundColor = 0x666666;
            player.skinBackgroundAlpha = 0.5;
        }
		
		public static function scale(target:*,width:int, height:int, shrinkable:Boolean = false, centered:Boolean = true):*{
			
			var dy:Number;
			var dx:Number;
			var ax:Number = target.scaleX;
			var ay:Number = target.scaleY;
			var ar:Number = ax / ay ;			

			target.width = width;
			target.height = height;

			var bx:Number = target.scaleX;
			var by:Number = target.scaleY;
			var br:Number = bx / by ;
			
			if(ar > br){ // la imagen debe ser mas chata, debo achicar scaleY
				dy = (ay * bx) / ax;
				target.scaleY = dy; 
			} else { // la imagen debe ser mas alta, debo achicar scaleX
				dx = (ax * by) / ay;
				target.scaleX = dx;
			}

			// control shrinkage;
			if ( shrinkable ){ 
				if ( target.scaleX < 1 || target.scaleY < 1 ) {
					target.scaleX = target.scaleY = 1;
				}
			}

			// center image;
			if ( centered ){
				target.x = width * .5 - target.width * .5;
				target.y = height * .5 - target.height * .5;
			}
			
			return target;
		}	
		
    }

}