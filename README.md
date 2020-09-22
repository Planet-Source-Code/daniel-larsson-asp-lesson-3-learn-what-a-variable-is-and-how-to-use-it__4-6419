<div align="center">

## ASP Lesson \#3 \- Learn What A Variable Is And How To Use It


</div>

### Description

In this lesson the actual ASP learning starts. My past two lesson was written so you can run ASP on your computer, and also a brief introduction to what ASP is and when to use it.<p>Here I will introduce you to the most basics of ASP and ALL script languages.....Variables. If you don't know how to use this, you won't get far in your Coding.<p>It might be intresting to read for you guys that aren't newbies, since I also explain the differens between Constants and Variables, and also very short about Arrays and Functions.<p>If you've read it, please rate it or comment so I know what im doin right/wrong when writing these guides.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Larsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-larsson.md)
**Level**          |Beginner
**User Rating**    |4.4 (44 globes from 10 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Documents/ Frames](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/documents-frames__4-27.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-larsson-asp-lesson-3-learn-what-a-variable-is-and-how-to-use-it__4-6419/archive/master.zip)





### Source Code

<hr><b>ASP Lesson #3 - Learn What Variables Are<p>At last! Now we will be looking into some real ASP Code. In this third
lesson I will teach you how Variables work and what a Constant is. I will also introduce you to Dynamic Variables and a little
function called Array. I will also teach you how to show the content of the variables in the webbrowser.</b><br><hr><br><b>Variables</b>
<p>Let us first understand what a variable is since variables are things you will use in almost every script you will write
in the future. A variable is, very simple, a name of a value. This value don't have to be determined directly, instead you
can determine this after the user/visitor has done something at your site that you asked him to do. Similar, you can determine
that the value is being calcuated by a formula that you choose.<p>
The reason to why we don't determine the variables at once, is that we don't know which value the variable may get from your
visitor. Similar, it's easier to work with a varible if you name it to something you can remeber, not just x,y,z,n etc. More about
this later in this lesson.<p>A variable can't just be only one value, it can also be a generic name for something that shall be
executed.In an exempel, a formula that must be run at more than one point in the script. Then it's unnecessary and clumsy that you
re-write this formula at all points you wish to add your formula. Instead, use a variable and name it to Formula1 or whatever you
wish to name it. To do this, just tell you script that this Variable is Equal to this formula.<p>Think about all the examples in
this lesson might be redicilous and that you might not find any useful ways to use this scripts! But variables are something <b>VERY</b>
basic for ASP, and even if we might be doing some redicilous ( sorry if i spell this word wrong, tell me how it should be spelled and I fix it )
things rigth now, you must understand this clearly before we proceed. Else it will only be hard for you in the future lessons.<p>
Ok hope I have explained well for you what a variable is, and why we almost use it in every script. In the next part I will teach
you how to write the code when you use a variable.<br><hr><p><b>Dim</b><p>
A script, almost, allways contains more variables, there are only few that uses one. So the first thing you must do when you write
ASP Code is that you must declare which variables the script shall contain and use.Example is below:<p>Dim,Tal,Mer,Siffra,Presentera<p>
With these lined i have Defined four variabled. Tal, Mer, Siffra and Presentera. These for variables have named that I have
figured out which means that they don't exist in own names in the VBScript Language. But the first word, Dim, is the word you
must use everytime you would like to define a variable.<p>When you are about to define more than one variable you use a comma(,) between
every variable to separate them from each other. Remeber that if you don't to this, two variables can be misstaken as one, or there will
be an error in the script.<p>Now you know how to define variables in VBScript, but now I'll teach you how to allocate values
to the variables.<p>Tal=20<br>Mer=20<p>What I just did was: Give the two variables Tal and Mer different values. Everwhere you
use the variables Mer and Tal in your script it will be treated as 10 and 20. At the end of this lesson I will teach and show
you how to use the variables together.<br><hr><br><b>Const</b><p>
Const means Constant. This works in the sam way as a Variable with one exception. The value that we allocate the const will
allways be constant. In other words, you can't vary this value.<p>Once you have allocated a value to a const, it will allways
be that value until you change it in the line/row that you defined the const. A const is defined at the same way as a Variable, with
one exception. You bothe declare and define a const at the same row.<p>Const Word="I like apples"<p>As you can see, the only
difference is that a const is defined and declared at the same row. If you use Dim to define variables, you must declare them at a
separate row.<br><hr><br><b>Array</b><p>I will only show you some very basics about what array is since I will be writing much
about arrays later in this guide. As you know Variables usually contains some sort of a value, but maybe you like it to contain
many values at the same time?<p>Dim Cars<br>Cars=Array("Saab","Volvo","BMW","Merceders")<p>Now the variable named Cars have four
different values. As you can see you place the different values within Paranteeses (), quotemarks and separates them with a
comma (,). Note: All variables that are allocated values that isn't clean numbers (1,2,3 etc) must be written inside quotemarks ("variable")
and be separated by a comma.<p>The reason why we use a parenteese () around these values depends, in this case, that we use
the function Array ( I will talk about functions in another lesson ).<br><hr><br><b>ReDim</b><p>You use ReDim whenever you want to
allocate a dynamic generic variable.<p>Dim Cars()<p>This indicates that the variable Cars still doesn't have any values, but from
this moment on exists and is ready to accept values. If you would like to allocate the variable Cars more values you use ReDim.<p>
ReDim Cars(300)<br><hr><br><b>Present Variables' Values</b><p>Okey, now you know what Dim, Const, Array and ReDim is. Now I will teach you
how a simple script with a few variables will be written. I will also teach you how to present them in your webbrowser. Example:<p>
&lt;html&gt;<br>&lt;%<br>Dim Number,More,Big,Total<p>Number=5<br>More=25<br>Big=1000<p>Total=Number+More+Big<br>%&gt;<p>The
total amount of these numbers is: &lt;b&gt;&lt;%=Total%&gt;&lt;/b&gt;.<p>&lt;/html&gt;<p>
Okey. What happened here? Since this is a complete script that will be rendered in the webbrowser we must first start with
"Blockers" or Delimiters that tells the webbrowser that this is some ASP Code. Then we will write our script. In this example,
the script defined and declared four variables and printed the total amount of these variables on the webpage. We also gave the
four variables different values.<p>The last variable, Total is the amount of the variables Number,More and Big. When the variables
have digital values (0-9) you can simply add them together using a plussign (+) between them.<p>Now the script is ready and we finish
the script by using a reverse block (%&gt;). The computer now knows what the answer will be, but the code will not show anything in the
webbrowser. To do this, we must ad some simple html code. First we write a sentance, The total amount of these numbers is, if we don't write
anything here a useless number will be displayed.<p>Since we wan't to show the variable Total we mark it a little extra with a B tag (&lt;b&gt;).
To finish, I write the variable Total within to Blockers, note: I also add a = sign before I write total. This means that the answer will be
printed here.<p>You show a variables value by adding a = sign.<p>&lt;%=Total%&gt;<p>Okey. You can now create a file in your ASP folder and run the script in your webbrowser. If you've done it right it should show the total amount of the three numbers we declared and defined =). If you have done something wrong, then check your code again. If it shows the correct amount in the page then view the source and look. Just look =). All the ASP code will be gone, and only simple HTML code will be there. Don't remember that the file must be named as *.asp!<p><b>End Of Lesson #3</b><br><hr>

