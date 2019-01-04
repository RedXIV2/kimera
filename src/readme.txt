 Kimera: Simple FF7PC 3D model editor v0.97a:
=============================================
	What's this?:
	-------------
		Kimera is a small characters/3D models editor for Final Fantasy VII (PC version). Features:
			-Display most of the 3D models found in the game. Field models and battle models including their 
			animations. The minigames models should work as well.
			-Supports editing the individual pieces that define the models (P files). Specifically, it can move 
			vertices, change their color, erase, create and cut triangles. It can also erase groups present on 
			a P file or change their rendering settings. Standard geometrical transformations 
			(rotate, resize and reposition) are also supported.
			-Add/Change/Remove textures (supports TEX, BMP, JPG and GIF files).
			-Add pre-cooked lighting for a single part and for the whole model (with up to four ligth sources)
			-Edit the bone lengths of a skeleton and attach and remove pieces to them.
			-Animations interpolation.
			-Field and battle animations can be edited, but the interface is rather unfriendly. 

	New in this version:
	--------------------
		-Fixed invisible semi-transparent models in-game.
		-Fixed group texture identifier edition.
		-Fixed lighting for battle models.
		-Fixed camera not resetting when openning the P Model editor.

	Usage:
	------
		-Skeleton view: The first thing you see when you open the program. Here you can:
			-Open a FF7 3D model.
			-Save.
			-Change the animation (only when you have loaded an HRC file).
			-Change battle animation (only when you have loaded a battle skeleton file).
			-Select and Resize/Rotate/Repostion the P models that compose the whole
			Skeleton (when aplicable).
			-Select and resize a whole bone or change it's length.
			-Add new P models to a skeleton bone.
			-Remove a P model from a skeleton bone.
			-Apply precalculated lighting.
			-Add/Change/Remove/Reorder textures for the model/part.
			-Edit the current frame of the current animation.
			SPECIAL KEYS:
				-Left mouse button: Rotate.
				-Right mouse button: Zoom in/out.
				-Both mouse buttons: Pan.
				-Home: Reset Rotation/Zooming/Panning.
				-CTRL+Z: Undo
				-CTRL+Y: Redo
				-Double click (in a P model): Jump to the P editor.

		-P Editor: Here you can change the P models detailedly:
			-Open a new P file.
			-Save the P file
			-Apply the changes in the loaded skeleton.
			-Resize/Rotate/Repostion.
			-Group trasparency level.
			-Several misc functions.
			-Rise or lower the brightness level of the model.
			-Assign a function to the left, right and middle mouse buttons. Left-click, right-click or middle-click 
			in one of the available functions (Bottom right corner). The one hi-lighted in red is assigned to the left button
			, the one hi-lighted in blue is assigned to the right button and the one hi-lighted in green is assigned to the 
			middle button.
			Available functions:
				-Paint: Click on a polygon to paint it with the selected color. If you hold shift while clicking
				on the polygon, you will select the color of that polygon.
				-Cut Edge: Cuts an edge. Thus, it divides all the adjacent polygons.
				-Erase: Click on a polygon to erase it.
				-Pick vertex: Click on a vertex and drag it where you want to move it.
				-New Polygon: Click on 3 vertices to create a new triangle.
				-Pan: Pan the model.
				-Zoom: Change the zoom level.
				-Rotate: Rotate the object (inspection).
			-Color modes:
				-Palletized: The old and imperfect method. It detects the model colors and let you change them
				(thus changing all the vertices with that color at once). The agressiveness of the detection
				is defined by the Threshold. I STRONGLY advise you aginst changing the Threshold once you have
				started to mess with the colors.
				-Direct: You can choose any color and then any of its gradations in the lower part of the pallete.
			SPECIAL KEYS:
				-Home: Reset Rotation/Zooming/Panning.
				-CTRL+Z: Undo
				-CTRL+Y: Redo

	History:
	--------
		v0.1:
			-Initial release
		v0.2:
			-Added Reposition/Rotation model suport
			-Added temporal Rotation (just to inspect the model). Can be done by pressing the mouse button
			inside the box with the picture and then dragging it.
			-Added the hability to erase all the polygons with a color diferent from the one selected and the
			complementary operation.
			-Hopefully, no more overflow problems with the battle models.
			-Less bugs in the color detection.
			-Now you can chose where do you want to save the changes.
			-New name :-P
		v0.3:
			-Added brightness control.
			-Added thin/fat level modifier (relative to the black line).
			-Added simple (maybe too simple) painting tools.
			-Added the ability to disable the precalculated lighting
			found in battle models.
			-Added Shift option. Turns a model without rotating it,
			useful for recycling non-simetric limbs.
			-Now the program won't try to open a non-p file.
			-Now the program can save the normals information. This
			enables the resulting files to be used with field models
			(even if they were for battle models originally).
			-Less unhandled errors?
		v0.3b:
			-New, more compact look.
			-Minor fixes.
		v0.4:
			-Better normals calculation. You can still chose to use the old method by checking "Calculate normals
			based only on polygons". This method is usefull when you want to preserve little details such as the
			buttons in Palmer's torso.
			-Added Paint polygon mode. When it's ON, you can't rotate the model, but you can change the color of
			the polygons.
			-Now the textured groups are allways white-colored (this way, their color doesn't interfer with the
			texture).
		v0.5:
			-Added New polygon mode. When it's ON, you can't rotate the model, but you can define new poligons by
			clicking on 3 existing vertices.
			-Added "Kill textured groups". Using this you can remove the textures from a model (useful por taking
			hairs from models, for example).
			-Removed the painting tools. They are no longer useful.
			-Improved the polygon selection method. Now it should be perfect.
			-Fixed several internal problems.
		v0.5b:
			-Solved the file corruption problems when adding polygons to multi-group
			(usually textured) models.
			-Solved (hopefully) the vertex sorting problem which made some of the added polygons
			invisible in-game.
			-Fixed the "Type mismatch" error when saving. Keep in mind that there must
			be still some kind of bug in the saving routine. Please, report any problems
			you find (provide the error number if possible).
		v0.5c:
			-Solved many GFX glitches when editing multigroup files.
			-Solved the dark textures problems and better normals calculation for
			multigroup files.
			-Now the unused vertices are phisically removed from the file
			(this was needed because FF7 seems to crash when it find unused vertices
			in a multigroup file).
			-Added the "Kill polygon" mode (a much more direct method for removing
			polygons).
		v0.6d:
			-Now the program uses OpenGL. Defintly looks better and is easier.
			-Added lighting support. Now you can set precalculated lighting like that one 
			originally aplied to the battle models. You can disable it, though, if you don't like it.
			-Added Zomming. Just press the right mouse button and move down to set the camera
			further or up to set it closer.
			-Added overwriting prompt. Now if you try to save your work in a file that already exists
			you will be asked to continue or not.
			-I had to redo the fat/thin control from scratch. Now it behaves in a more precise way.
			(not necesarily a good thing...)
			-Killed many bugs (such as the wrong distance calculation that caused models to be 
			completly unviewable).
			-Killed many silly bugs releated with the normals
		v0.7:
			-New GUI: Now we are closer to the all-in-one model editor. You can view the whole models decribed by the
			HRC files, test animations, change the length of the bones, add more parts to a bone and, of course, edit
			the parts of the model by double-clicking on them. Now you can do everything directly from Kimera but working 
			with textures (patience, I'll probably end up adding it too).
			-Added General lighting. Now you won't need to set the light for every single part if you want to add
			pre-calculated lighting to a model.
			-You can now associate the HRC files to this program.
			-Several bug-fixes.
		v0.8:
			-Rewriten (nearly) from scratch.
			-Support for Battle Skeletons.
			-Multi-P bones are now compiled in a single P file.
			-Resize/Reposition/Rotate directly in the skeleton view.
			-Paning in both the Skeleton view and the P editor.
			-Now you can chose between the classic (and buggy) palletized colors and the direct color control 
			(simple and more acurate).
			-You can now assign independently functions to the right and left mouse buttons in the P editor.
			-New functions on the P editor: Cut edge and move vertex.
			-Killed many bugs, but, since I've rewiten nearly all the code, I'm quiet sure there will be new ones.
		v0.81b:
			-Added basic support for battle locations.
			-Now textures are displayed.
			-Killed many bugs and glitches.
		v0.82:
			-Optimized saving/loading routines and 3D Rendering (using display lists).
			-Corrected texture glitches.
			-Groups properties (only trasparency level, right now).
			-Skeleton bones visualization.
			-Added a default animation (resembling "The man of Vitruvio" pose) used when
			no animation could be loaded.
			-Killed many bugs here and there.
		v.083:
			-Killed the fatal bug that was causing heavy data corruption with some models 
			(or so I hope).
			-Solved a few texture problems.
			-Added a rudimentary (and experimental) animation edition tool. Only for field
			models currently.
			Keep in mind that now also the animation will be written to disk when you save 
			your model.
		v.084:
			-Added battle aniamtion playback. Still imperfect. Some animations may show wrong
			rotations and not every animation can be loaded.
			-Added summons support (maybe other magic.lgp models?).
			-Now you can also add pieces to battle models. Keep in mind that they will be joiend 
			when you save the model.
			-Added a button to remove pieces (yeah, I should have added it long ago...).
			-D-Lists management has poven to be buggy, so I added an option to disable this 
			optimization.
			-Solved most of the texture clipping problems.
			-Killed several minor bugs.
		v0.9:
			-Added support for weapons and their animations.
			-Added support for battle animations edition. The whole animations pack will be saved 
			along with the model.
			-Added a textures loader (can load TEX, BMP, JPG, GIF and ICO files). Textures will be 
			saved as TEX files along with the rest of the model.
			-Updated the field animations and TEX formats so they match the lastest specification 
			posted on the wiki.
			-Battle animations are now decoded properly.
			-Solved the precision problems when editing animations.
			-The usual round of minor bugfixes.
		v0.9a:
			-Bug fix (including the one which was causing edited models to crash the game).
		v0.9b:
			-Improved field animation detection.
			-Added a checkbox to set the transparent color flag for textures.
			-Bugfix on geometry edition.
		v0.91:
			-Killed several bugs releated to the textures.
			-Solved some clipping issues.
			-Added an option to see the ground (with the circular shadow of the model projected on it).
			-Added an option to propagate animation changes to the following frames (Propagate f.).
			-Added an option to see the "ghost" of the last frame.
		v0.91a:
			-Solved a bug that repositioned battle model weapons while saving.
			-Solved a bug saving summons skeletons.
			-Added an option to make the animation set the model above the ground.
			-Added an option to synchronize weapons position to that of a specified bone.
		v0.92:
			-Killed several bugs.
			-Fixed the "Delete all polygons but those with the selected color"/
			"Delete all polygons with the selected color". Well, actually their code was simply empty since v0.8.
			-Added undo/redo buffer. You know how this works: CTL+Z = Undo, CRTL+Y = Redo	
		v0.93
			-Bugfixes. A lot of them:
				-The produced animations now work correctly in-game
				-The textures are now correctly updated.
				-Several vertex/polygon picking glitches.
				-Jut too many more...
			-Added the option to use point ligths (the way thing should have originally been...)
			-Added an option to show the axes on the P editor.
			-Added an external configuration file (kimera.cfg). Only the Undo/Redo buffer length can be set up.
			Setting it to 0 will disable it.
		v0.95
			-Added 3Ds loading support.
			-Added several operations that depend on a given plane:
				-Cut the whole model through a plane.
				-Make the model symetric.
				-Erase everything under the plane.
				-Mirror on a given plane (as oposed to the fixed ZY plane used in earlier version).
				-Fatten/thin on a given plane (as oposed to the fixed XY plane used in earlier version).
			-Adjusted the model rotation to fully use quaternion properly, no more gimbal lock. The animation rotations are still
			the same, though.
			-Now you can hide groups on a P model. When a group is hidden it can't be affected by any operations on the editor
			besides the palletized ones, rotation/translation/panning and the group deletion.
			-Killed lots of bugs (you'd think there couldn't be many more after all this time... well, there were).
		v0.95c
			-Added render state edition and preview.
			-Fixed blending modes preview.
			-Fixed the missing textures in-game bug
			-Fixed the bug on RSB file names generation.
		v0.96
			-Added animation blending for field and battle skeletons:
				-For a single frame: Adds a new frame between the current and the previous. Assumes the animation is a loop (and thus, the 
				previous of the first is the last)
				-For the whole animation: Doubles the number of frames. If the animation is not a loop the firt->last frame is discarded.
			-Fixed some overflow errors.
		v0.96b
			-Added an option to specify the number of frames that should be added between frames when interpolating.
			-Fixes a few crashes releated to un-do/re-do when changing animation length.
		v0.97
			-Added support for limit break animations (needs data from battle.lgp).
			-Added a dialog to load animations that are used by a field model at some point (using Ifalna's filter)
			-Added a dialog to interpolate all the animations in char.lgp, battle.lgp and magic.lgp.
			-Implemented full compression support for battle animation packs.
			-Fixed a ton of bugs releated to animations interpolation and writting.

	Known bugs/Coments:
	-------------------
		-The color detection routine it's still far from perfect (and most likely will allways be).
		-D-List management is buggy. If you start to see screwed or repeated pieces or they simply vanish
		try disabling this optimization.
		-The normals calculation during most editing operation. If you find that your model becomes strangely
		shaded while editing it, disable lighting and enable it again to force a full recalulation.

	TODO list:
	----------
		-Add mesh-smoothig (Half done).
		-Helper to move vertices and joints with ease.
		-Per vertex ambient occlussion.


	Thanks:
	-------
		-Square soft, for creating one of my favourite games.
		-Eidos, for porting it to the PC.
		-Reunion, for beginning the battle->field project conversion.
		-Mirex, for his excelent tool Biturn.
		-Ficedula, for his field animations filter (from Ifalna)
		-L.Spiro & Qhimm, for their great documentation about the Battle animation format.
		-Ahlexx, for his great documentation about the P and A files format.
		-Aali for his notes about the TEX and A file formats.
		-seb, for his notes about the weapon animations.
		-Qhimm, for his great forum.
		-Kim Shoemaker, for his excelent generic euler<->quaternion code.
		-Larry Rebich, for his code to deal with the browse folder dialog. 
 		-And the people from Qhimm's forum for betatesting and brainstorming.

	Disclaimer:
	-----------
		-Use this code as you wish but please, give me credit for my work. It hasn't been easy for me to write this program.