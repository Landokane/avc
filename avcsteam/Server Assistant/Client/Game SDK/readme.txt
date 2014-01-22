Avatar-X's Server Assistant Client Game Creation Software Development Kit (ASACGCSDK)
-------------------------------------------------------------------------------------

This little program will "emulate" playing a SAC game thru SAC over the net. 
You can start by editing frmGame2.frm. In order for it to be loaded, the form must be called frmGame2.

A few notes:
============
- Don't actually REFERENCE the form, like frmGame2.Show or frmGame2.Left = 30. Since the ACTUAL form isn't loaded, (only an instance is) this will cause problems.
- A lag slider has been included. If 2 seconds of lag isnt enough, just edit the scroller properties.
- Don't mess around in the BAS file, changing code there to suit your game will not work.
- You can use the IsOpponent variable to figure out who is who (IsOpponent is FALSE on the instance that STARTED the game, or in this case, on the one on the left)
- Keep the amount of packets and the size of the packets DOWN! This all has to go through SERVER ASSISTANT, and the main priority for SA should be the game server itself.

How SENDIT works:

When you SENDIT on Game A, it will arrive in Game B in the GameInterprit code. 
The CODE is meant to be a 2 letter all caps code.
The PARAMS (or DATA) can be almost anything, except characters over 239.

Example:

SendIt "AB", "Some Text."

In GameInterprit, this will arrive as the variable a$ set to "AB", and the variable p$ set to "Some Text".
Then you would put something like:

If a$ = "AB" then
    MsgBox "The other guy says " + p$
End If

or whatever it is you are doing.

If you need to know any more, just let me know!

Avatar-X
avatar-x@cyberwyre.com





