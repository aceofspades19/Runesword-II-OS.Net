�    Tester   The Northlands is a place of rangers, hunters, and trackers--masters of the forest. Northlanders have a reputation of being untrustworthy and eccentric.  Of course, they are hired still when Royalty go on hunting expeditions. The Northlands is run by a council of Elders.
       Eternia\Homeless.tomI 8   ranger1.bmpC                       
         2 ( (                            
   Gold Piece     d d              
                  	   gold1.bmp             +2 Wooden Staff <   This staff is of excellent craftsmanship. light and elegant. �        
@              
   staff2.bmp                Fencing �   Using a sword, you may disarm attackers when they miss with a melee weapon.  At Novice rating, the chance is 20%, and it increases 5% per Skill Level.,
   
  /    6          make sure I'm using a sword 6          (I'm the target in this case) 6       +   (because this trig fires when I'm attacked) 6            	        sword 	          4	       <       	 5       
             
            6            6       -   Make sure my attacker is using a melee weapon 6            	                  <                         
            6            6           exit if it's not sword vs. sword 6                                           6            6       +   Also exit if attack succeeded in hitting me 6            _                               ! 6           " 6           # 6       I   chance for disarming is 20% for novice and +5 per skill rating from there$ 6           % 5  c      &     '     ( 6           )       *            +            , 6           - 6          Let's drop the attacker's sword. 6           /                           Disarm     <           	          4	       <        :          $      /   DM|* Fencing Skill disarms [CreatureNow.Name] *                         
                            Stop at Level 10     >            	       $       9   DM|You cannot raise this Skill any higher than Master +6. 5                                     