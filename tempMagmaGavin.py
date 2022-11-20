const rightMost=72;
const leftMost=8;
const topMost=36;
const bottomMost=8;

hero.moveRight(1);
while(true){
       if(hero.direction=='up' && hero.y >= topMost){
           hero.fire();
           hero.turn("left");
           }
       if(hero.direction=='right' && hero.x >= rightMost){
        hero.turn("up");
        }
        
   if(hero.direction=='left' && hero.x <= leftMost){
       hero.turn("down");
       }
  if(hero.direction=='down' && hero.y <= bottomMost ){
       hero.turn("right");
        }
}
