Dynamic Technologies Graphical ScrollBar v1.0
  - A replacement for the Generic / Limited / Basic VB ScrollBar but with Soooo much more.

(c) 2000
Howard D. Hull Jr.
Dynamic Technologies 

'==============================================================================
'This code is copyrighted and has no warranties either expressed or implied.
'==============================================================================
'Terms of Agreement:   
'By using this code, you agree to the following terms...   
' 1) You MAY use this code in your OWN programs (and may compile it into a program and distribute it in compiled format for langauges that allow it) freely and with no charge.
' 2) You MAY NOT resell this code / control as it's own entity. 
' 3) You MAY NOT redistribute this code (for example to a web site) without written permission from the original author. Failure to do so is a violation of copyright laws.
' 4) You MUST give proper credit to Howard D. Hull Jr. and Dynamic Technologies as the author of this code.
'==============================================================================

Orientation is automatically determined when you first Draw the control on the form. 
  Draw a tall control for a Vertical ScrollBar or long for a Horizontal Bar.

The scrollbar uses pictures for the Background, Left/Right/Top/Bottom Buttons and the Thumb. 
  Optionally the Buttons can be hidden with the ButtonsVisible property set to False. 
  This will let you scroll from one end to the other. 

ButtonsVisible = False
     +---------------------------------------------------+
     |                                                   |
     +---------------------------------------------------+
     ^           We can Scroll the entire length         ^

ButtonsVisible = True
     +----+-----------------------------------------+----+
     |    |                                         |    |
     +----+-----------------------------------------+----+
          ^  We can only Scroll between the buttons ^

The Thumb and Button pictures can be aligned based on the orientation of the ScrollBar. 
  When Vertical : Left/Right/Center
  When Horizontal : Top/Bottom/Middle
     +-----+---+-----------------------------------------+
     |     |   |                                         |
     |     +---+       +---+                             |
     |                 |   |                             |
     |                 +---+                +---+        |
     |                                      |   |        |
     +--------------------------------------+---+--------+
             ^           ^                    ^
            TOP        MIDDLE              BOTTOM

Let's you tweak the position of the thumb on the background. 
This property will also affect the alignment of the Buttons.

Setting the AutoSize property to True will resize the control to the Height/Width of the BackGround picture.

Min and Max properties are Longs... No more 65,535 limit. 
  The only restriction is the difference between the Min and Max values has to be less than 2,147,483,647. 
  Negitive values are supported. 

Scroll & Change Events work exactly the same way as the basic ScrollBar.

Value Property ... Self explanitory.



 - I am currently developing version 2.0 which will hopefully encorporate the following enhancements;
	+ Clicking and holding down the mouse on a button or the background, will keep scrolling.
	+ MouseDown Pictures for the buttons and the Thumb.
	+ Incorporate Image Masks to allow the Background image to be oddly shaped.
	+ Different Border styles, Recessed, 3D ... Mainly for bars without a BG picture.
	+ Incorporating into a Grphical Listbox
	+ Creating other Graphical Controls, CheckBox, OptionButton... etc.
	+ Addition of MouseDown, MouseMove, MouseUp events (If I see a usefulness. Let me know.)

- If you have any comments / complaints / praise :) let me know [ howard@dynamic-technologies.net ]