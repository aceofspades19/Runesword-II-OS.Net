�    Urane Urelia0 �   Kyparissia is the land of scientists and explorers, hard-headed men and women who are legendary for their rudeness, or "blunt honesty" if you are kind. They feud with the Argosans, who are of dissimilar temperment.       Eternia/Home/Home.tomK @   wizard3.bmp                    
     p    # #                            
   Gold Piece                  1J                 	   gold1.bmp             Ring of Death &   This ring makes your finger feel numb. �           �IB                 	   ring2.bmp        Lightning and Runes �   Snitched from Magestorm, edited a bit, some from Ice Storm.  Eats up a little bit of health to hurt enemies every combat round, but will not drain a character below 1/2 (round up);        '    6          If not in combat, then exit                                  6            6            6       )   must have at leat 1/2 of your max health.     	     
                                6            6           Like anyone would have less than 6          6 health after the last check? 6           Could happen, if their max was 3 6          or something.                                     6                   -           6          Don't shoot the bearer! � �       N N                              $      *   DM|[ItemNow.Name] strikes [CreatureA.Name] 6          PlaySound 'Zap11.wav'  .      	   smoke.bmp!     " 
       #     $ 
           % 6           & 6           ' 6                                Curved Sabre       K         !JB         	      	   sword.BMP             Wand of Sparks 1   This wand is inlaid with gold and silver tracings   
           6J                 	   wand2.bmp        Sparks CreatureNow     "           5             $       p   DM|The wand explodes in a blazing shower of sparks and fire!  [CreatureNow.Name] suffers [Local.IntegerA] damage                        Athleticism ^   Bonus to ActionPoints several times per day.  Always gives same bonus to swim and climb rolls.'        #  
        6 �         7            9 6            5   
       5        1 6           8 �         , 6           4     : 6                  3           $       /   DM|[CreatureNow.Name] whirls about in a frenzy!     ; 6                        $       Z   DM|You can only use Athleticism [Local.IntegerA] times per day.  You've reached the limit.< 6                                     	   Lunaspell i   This skill alows you to cast spells of Lunacy. Hallmarks include loss of senses, necromancy, and madness.'        6       <   For a caster with Will 10, it's 50%. For every point of Will 6       4   above, subtract 5%. Conversely, penalize dummies +5% 6          Max is 95% Failure, Min is 5%. 6            2
           ; �       5;                    	            
 6            
`          2     6            _       5 _              5                    6            /                         Bravado Delusion }   Targets believe they are invulnerable for 4 turns. They waste time with posing and short range attacks get a +6 against them.)                              .         starspin.bmp !       	   Magic.wav                     $      ,   dm|Bravado delusion strikes [CreatureA.Name]                +6 against if short range     +     �    e       ��                                 Posing     ;     �                                          7     $      B   DM|[CreatureNow.Name] strikes a few attack poses, absurdly pleased                 Give Appropriate Spells     ?                                $          DM|You gained 2 new spells!                                 $          DM|You gained 2 new spells!           
                      $          DM|You gained 1 new spell!            5                                   Shadowguard �   Party members within 2 spaces of Caster get a -2 to Defense. Any monster coming within 2 spaces of Caster must save vs. will or run in fear. Lasts 2 turns per Skill Level of Caster.)                  6       .   Set the duration of triggers to Caster's Level 	               5 N       
            6          Now copy triggers 	           	          
 
            	 
                    e       $      *   |Whispy shadows surround [CreatureA.Name].                       
                           -2 to Defense     ;     @    ��                     In Range Scare    Caster's Name goes here;     @    6          Set Target equal to Caster 	 
          N                                         
            6       2   If close, scare 'em and set this trigger to remove	 e      
 `        $      +   |The shadows howl around [CreatureNow.Name]            5                    $      .   |[CreatureNow.Name] is unafraid of the shadows                                        Permanent Scare     ;            5�                        Finger of Death U   Kills one short range target. This spell does not work against animals or the undead.)      @    
              e       $         DM|Too Far!                         .        bolt.bmp !         Zap9.wav	 $       3   DM|Finger of Death will be cast on [CreatureA.Name]
 <        $       @   DM|[CreatureA.Name] is undead and unaffected by Finger of Death!                         .        $      C   DM|[CreatureA.Name] is an animal and unaffected by Finger of Death.                         6            $      +   DM|[CreatureA.Name] collapses to the ground 5                      	   Ice Storm S   Casts Caster's Level number of icicles doing 2d8 in damage to random enemy targets.)                   !      	   Zap10.wav 6          [Titi] count how many foes 	                                    
           	       
 -           .        bluemist.bmp 5	 N       $      &   DM|[Local.TextA] feels an icy blast...              6       ?   [Titi] Each time one enemy dies, decrease the number of targets                     6       G   [Titi] and when there are none left, no need to continue the ice storm!                                     
                            Stricken Soul y   Target gets -5 on to hit rolls or misses turn. Also has a chance of causing the target to drop his weapon. Lasts 3 turns.)      @                      6                                6            !       	   Zap11.wav .    	   ball3.bmp	 $      )   dm|Stricken Soul strikes [CreatureA.Name]
 6                  c       	          <        5                               
                    :         $       ?   DM|[CreatureA.Name] screams in pain, dropping its [ItemA.Name].                                        Freeze     ;     �             5d                                 	   -5 to hit     ;     �    ��                     Frenzied Mortis �   Brings dead creature back to life in an agitated state for several turns before lapsing back into death.  Usually (90%) of the time, the creature allies itself with you.)      @!    	      6       4   This crazy spell animates a dead guy for a few turns         .       bolt.bmp !       
   Magic7.wav 6                    $      -   dm|Frezied Mortis only works against the dead	            
             6            $      +   dm|Frenzied Mortis strikes [CreatureA.Name] 5        5        6            6       :   Pump up the Creature's stats based on Caster's Skill Level 6                 [[     OO     ``     6            6       >   10% chance dead guy joins wrong team (feeeeeelin' evil tonite) 6             	        $      0   DM|[CreatureNow.Name] senses something amiss.... 5                     5                    6                      !                       
   Slow Death     <
     �           5                                     Energy Boost     ;      @         5                        Thought Loss	 g   All enemies lose 1d6+3 Will, Agility, and Strength a number of turns equal to the Caster's Skill Level.)       
           	               
            6            	            .    
   wapple.BMP !       
   Magic9.wav	          
 
                           sap saw     ;     @    �P�     �P�     �P�     $      .   DM|[CreatureNow.Name] is woozy and disoriented .        starspin.bmp !       
   Magic3.wav                 no helm no shield
                  	        shield 
        helm 6            	          6                    6            	       $       =   DM|Spellcasting is not possible while equipped with a shield. 5; �       
       $       ;   DM|Spellcasting is not possible while equipped with a helm. 5; �                   6           	             6           
 
            6                            Wizardry n   Your targeted hostile spells are +5% harder per skill level to save against for every opponent in this combat.'    �     6           %  
        ' �         (            ) 5   
      * 5        & 6           + �         - 6           9       >         : $       C   dm|Wizardry is a combat skill that requires 5 action points to use.;            <            = 6           0     /       1 $       0   dm|You cannot use Wizardry again until tomorrow.2            8            . 6           4     ! 	         " 5        #            $ 6            2                          Wizardry Effect 7   Knock down all vices, giving a -X% to saves vs. spells.             	           	                   P P     R R     a a     
                            Stop at Level 4     >                   $       6   DM|You cannot raise this Skill any higher than Master. 5                                 
   Wrathspell i   This skill allows you to cast spells of Wrath. Hallmarks include elemental fire, direct damage, and rage.'        6       <   For a caster with Will 10, it's 50%. For every point of Will 6       4   above, subtract 5%. Conversely, penalize dummies +5% 6          Max is 95% Failure, Min is 5%. 6            2            ; �       5;         5H d      	            
             6            
`          2     6            _       5 _              5                    6            5H        /                         Hotmetal �   Heats up metal armor (chain or plate mail) to an unbearable level. Enemies suffers 4d8 damage and attempts to shed armor until it is off or cools down.)                             !          Fire.wav .    
   flame2.bmp           $      $   dm|Hotmetal strikes [CreatureA.Name]             	   Hot Armor     ;     @    	        chain 
        plate 6            	          6            4	       4
       6           	        
 $          DM|[ItemA.Name] glows orange-red                    5         5 [       $       J   DM|Agility Roll: Need a [Local.IntegerB] or more.
Rolls: [Local.IntegerA] 6                   :          $      5   DM|[CreatureNow.Name] successfully sheds [ItemA.Name]             $      )   DM|[CreatureNow.Name] fails to shed armor             6                        6                        6            
                            Give Appropriate Spells	     ?        $                        F        Unthinking Fury          $          DM|You gained Unthinking Fury!                        F        Jumpfire	         
 $          DM|You gained Jumpfire!                                  F     
   Desolation          $          DM|You gained Desolation!                        F     	   Fireblast          $          DM|You gained Fireblast!                                  F        Scorching Suns          $          DM|You gained Scorching Suns!                        F        Fey Darkness           $          DM|You gained Fey Darkness!!           "            # 5       $                            Scorching Suns '   Obliterates target in a column of fire.)           	    	   !       !       
   Magic2.wav $       /   DM|The air burns and crackles with fiery energy !          Fire.wav .    
   flame1.BMP
 �                         Jumpfire �   Target suffers 3d8 fire damage for 1 turn.  This blanket of flame gets its name from its habit of "jumping" randomly from target to target.)                                        !          Fire.wav .   
   flame1.BMP $      $   dm|JumpFire strikes [CreatureA.Name]                       jumpfire damage     ;     @    -           .    
   flame1.BMP !          Fire.wav $      <   DM|A fiery halo of flame bursts around [CreatureTarget.Name]                     	   Fireblast ]   Target suffers 4d10+4 fire damage.  This spell has a 3% chance of backfiring upon the caster.)       
                   6             c       $       -   DM|Fireblast backfires on [CreatureNow.Name]!                        !          Fire.wav	 .    
   flame1.BMP
 v                    
   Desolation ]   Drains 3d4 points of Strength and Agility each round for 1d6 turns. The effect is cumulative.)       	                       	          5        
            !          Magic13.wav .      icehoop.bmp          	 $      &   DM|Desolation strikes [CreatureA.Name]                Drain Strength and Agility     ;     �                � �     � �     $         DM|* Desolation *                 Unthinking Fury ~   Target gains +15 to energy and +15 to damage. Target may become irrational, however (10% chance each turn). Lasts 1d6+1 turns.)                             !          Zap7.wav .       bluemist.bmp 	          5        
                                     crazy stuff (bad)     ;     �         ��                          $       >   DM|[CreatureNow.Name] is smashing his head against the ground.                         	 $      '   DM|[CreatureNow.Name] regains composure
 6                      -           >            $      '   DM|[CreatureNow.Name] regains composure 6                      $       7   DM|[CreatureNow.Name] sits down and howls for a moment. !       	   Roar8.wav 5        6                      	          :          
            $       O   DM|[CreatureNow.Name] laughs hysterically and drops all equipment and supplies. 6                                        Fey Darkness     )            	        5         5         5         5         5         5         5        	 5        
 5         5         5         5         5         5         5         5         5         5         5         5                         Firefingers �   Combat:  Target suffers 8 points of fire damage per caster's level.  Range is 3 spaces. This spell has a 10% chance of backfiring upon the caster.)       !    6                     6            6                                6            $       '   DM|This spell has no target except you!	 $       /   DM|Firefingers backfires on [CreatureNow.Name]!
            !          Fire.wav .    
   flame1.BMP                                      6            6            6                     6             c
       $       /   DM|Firefingers backfires on [CreatureNow.Name]!                        6            !          Fire.wav .    
   flame1.BMP 6            6                  6           !                         no helm no shield                  	        shield 
        helm 6            	          6                    6            	      	 $       =   DM|Spellcasting is not possible while equipped with a shield.
 5; �       
       $       ;   DM|Spellcasting is not possible while equipped with a helm. 5; �                   6                        6            
            6                            Define home     E        
            Eternia/Home/Home.tom $      f   DM|[CreatureNow.Name] was born in [CreatureNow.Home]. Does [CreatureNow.PronounHeShe] still live here? %          Yes %          No &                  $      2   DM|Where does [CreatureNow.PronounHeShe] live now? 3	         	 5 	      
                             Phule's Safety Net !   Does the protection of the party.6        ,    6            6       !   if someone in the party is alive, 6          quit. 	 
                                          
           	 6           
 5        6          everyone's dead!?! - Fix it. 	 
          6            6       #   no special reason - just want to be 6          careful of negatives        5                    6            6          raise to health 5        
            6            6          now end the combat 	            	     9    	     5 d       5        	          5          :        ! 
           " 
           # 6           $ 6          now tell them what happened% $       A   DM|You thought you were a goner - but you seem to still be alive!& 6           ' 6       !   you will need to change this next( 6          line to suit your tome.) @     Arena* 6           + $       2   DM|You wake up in the first aid room of the Arena., 6                            Were    Generic Were trigger;      r    5         6            6          Shapeshift conditions 6             
                               6           	 6       3   What follows is the logic of the actual shapeshift.
 6            6       %   Is CreatureNow in party or encounter? 5         	
          N N       � �       5                    
            6            6            6            6            6       5   Set the creature comments in this trigger to the name 6          of the creature shapeshifting 5	        	           6            5 �       5�         5O O       5�         6           ! 5 �      " 5�        # 5[ [      $ 5�       % 6           & 5 �      ' 5�        ( 5` `      ) 5�       * 6           + 5a a      , 5P P      - 5R R      . 5       / 5       0 5       1 6           2 5       3 5       4 6           5 5 N      6 5N N      7 5S S      8 5       9 5       : 5       ; 6           < 5       =        > A        ?            @ A        A            B 
           C         D 6          Error condition.E            F            G        H 	 
         I N       J 5	       K 5       L          M 5        N            O 
           P            Q 	           R N       S 5	       T 5       U          V 5        W            X 
           Y            Z        [ 6          Error condition.\            ]            ^ 	         _ 5        ` :         a 
           b 	         c        d        e A         f            g A         h            i           j            k 
           l $         DM|[CreatureNow.Name] howls.m 5       n 6           o 6           p 6           q 6           r 6              �    Urane Urelia �   Kyparissia is the land of scientists and explorers, hard-headed men and women who are legendary for their rudeness, or "blunt honesty" if you are kind. They feud with the Argosans, who are of dissimilar temperment.   Were   Eternia\Homeless.tomK  	   Bunny.bmp*                       
     
     # #                                   Were Invulnerability 2 4   Avoid weapon attacks unless they are done by silver.,            BQ       Were                         	          B4       silver <                               	 
           
 _     $      *   DM|The attack passes through unhindered...                 Doesn't Fit                 6            $          DM|Nothing seems to fit.... 5                        Were    Generic Were trigger;       s    Q        Were 5         6            6          Shapeshift conditions 6                                           	 6           
 6       3   What follows is the logic of the actual shapeshift. 6            6       %   Is CreatureNow in party or encounter? 5         	
          N N       � �       5                    
            6            6            6            6            6       5   Set the creature comments in this trigger to the name 6          of the creature shapeshifting 5	        	           6            5 �       5�         5O O        5�       ! 6           " 5 �      # 5�        $ 5[ [      % 5�       & 6           ' 5 �      ( 5�        ) 5` `      * 5�       + 6           , 5a a      - 5P P      . 5R R      / 5       0 5       1 5       2 6           3 5       4 5       5 6           6 5 N      7 5N N      8 5S S      9 5       : 5       ; 5       < 6           = 5       >        ? A        @            A A        B            C 
           D         E 6          Error condition.F            G            H        I 	 
         J N       K 5	       L 5       M          N 5        O            P 
           Q            R 	           S N       T 5	       U 5       V          W 5        X            Y 
           Z            [        \ 6          Error condition.]            ^            _ 	         ` 5        a :         b 
           c 	         d        e        f A         g            h A         i            j           k            l 
           m $         DM|[CreatureNow.Name] howls.n 5       o 6           p 6           q 6           r 6           s 6                           Fail    Can't raise this 'skill'>            5                        Infect ,   Infects the target creature with lycanthropy2                                           	          B       Were                         
           	          
 5         	
          N N       � �       5                    
                     2                       2                                      Infect From Encounter    Sub trigger to get job done.2            	          B        Were A                               
            	           5       	 5       
 	          B        Were                       
            
            	                      
                            Infect from Party    Sub trigger to get job done.2            	          B        Were A                               
            	           5       	 5       
 	          B        Were                       
            
            	                      
            6                            Were    Generic Were trigger;      r    5         6            6          Shapeshift conditions 6             
                               6           	 6       3   What follows is the logic of the actual shapeshift.
 6            6       %   Is CreatureNow in party or encounter? 5         	
          N N       � �       5                    
            6            6            6            6            6       5   Set the creature comments in this trigger to the name 6          of the creature shapeshifting 5	        	           6            5 �       5�         5O O       5�         6           ! 5 �      " 5�        # 5[ [      $ 5�       % 6           & 5 �      ' 5�        ( 5` `      ) 5�       * 6           + 5a a      , 5P P      - 5R R      . 5       / 5       0 5       1 6           2 5       3 5       4 6           5 5 N      6 5N N      7 5S S      8 5       9 5       : 5       ; 6           < 5       =        > A        ?            @ A        A            B 
           C         D 6          Error condition.E            F            G        H 	 
         I N       J 5	       K 5       L          M 5        N            O 
           P            Q 	           R N       S 5	       T 5       U          V 5        W            X 
           Y            Z        [ 6          Error condition.\            ]            ^ 	         _ 5        ` :         a 
           b 	         c        d        e A         f            g A         h            i           j            k 
           l $         DM|[CreatureNow.Name] howls.m 5       n 6           o 6           p 6           q 6           r 6                           Fail    Can't raise this 'skill'>            5                        Regen }   This skill will regenerate a creature, even from death.  Skill level determines rate of regeneration.

Version 1.0 by Phule;
   2      F
        Regen                                               Regen 5   Does the regeneration for those with the Regen skill.            	 
          F        Regen                      5c                                	 
           
 	            F        Regen                      5c                     d       9    d                             
                            ClawBite     &        	           ?        Fang Attack                  ?         Claw Attack      ?         Claw Attack     	                             ClawBite     I        	           ?        Fang Attack                  ?         Claw Attack      ?         Claw Attack     	                                          Fail    Can't raise this 'skill'>            5                         