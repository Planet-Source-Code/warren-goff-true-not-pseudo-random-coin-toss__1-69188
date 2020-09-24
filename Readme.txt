As you know, computers generate a pseudo set of random numbers. This is unavoidable in a digital environment. A coin toss is in reality an analog event. The only way to have a truly random number generator is by introducing an analog to digital conversion. This program does exactly that. The analog random event sampled is sound frequency and by using a microphone, a completely random set of numbers (seeds) are generated. The audio FFT is borrowed from PSC. Many methods have been suggested to perform this task. Some have used random data sets as seeds. Others have sampled mouse movement. A website provides realtime random atmospheric data (www.random.org). There has even been a proposal to include a cesium chip in all computers eventually as it decays in a random fashion. My method seems foolproof as long as a microphone is set at a reasonable sensitivity and the environment is noisey. The importance of true randomization cannot be overemphasized as this is essential for the scientific method and in clinical studies. My application, which will be posted soon, is simpler and more abstruse. I intend to use this random generator to identify variations in our collective unconscious, an idea I heard about on National Geographic. 

---Update(08-22-07):Instability added, minor changes

The definition of a random number or randomness in general is quite vague:

"So, what is a random number? You will realize that this is a silly question as soon as you rephrase it in the form ``is 2 a random number?''. Randomness is not a property of individual numbers. More properly it is a property of infinite sequences of numbers. "..."As far as I am aware, nobody has ever given an entirely convincing definition of the term ``random sequence''. On the other hand, everybody has a common-sense idea of what it means. We say things like: ``the sequence should have no pattern or structure''. More directly we might say that knowing x1,..., xn tells us nothing about xn+1,...." [Knuth, D. E. (1981), Seminumerical Algorithms, Vol. 2 of The Art of Computer Programming, second edn, Addison-Wesley. ]

There are basically 2 kinds of random number generators (RNG), Pseudorandom number generators and Hardware random number generators. My project is the latter. Does this generate a truly infinite set of unpatterned random numbers? I guess that is impossible to definitely know. Are there better analog sources to seed the generator? It appears that there are although they are not readily available to me.

I am not a statistician unfortunately but there are methods such as Chi Squared and KOLMOGOROV-SMIRNOV used to assess the true randomness of a data set. The Visual Basic RNG is a linear congruential generator (first-order Markov process) and "A good linear congruential formula will generate a long sequence of numbers before repeating itself...The VMS Run-Time Library provides a Pseudorandom number generator routine called MTH$RANDOM...if the algorithm is repeatedly called, the value of SEED will alternate between EVEN and ODD values. MTH$RANDOM, RANDU, ANSI C, MICROSOFT C, etc all will fail.

For all practical purposes, it would seem that pseudorandom numbers are adequate for most circumstances such as simulations. "But what is the
difference between a good sequence of random numbers and a bad sequence of random numbers?...A good sequence of random numbers is one which makes us believe the process which created them is random." "If you only need a small number of random numbers, say less than 100,000, if you do not demand high resolution of the numbers, and if you're not concerned about correlations between the initial SEED value" then a linear congruential formula is adequate. If one adds "shuffling" to the algorithm, 50 million numbers can be generated randomly with high resolution.

My project attempts to be a Hardware random number generator, akin to actual tossing coins, drawing lots, rolling dice and blowing ping pong balls as in bingo. In a sense, it samples atmospheric chaotic noise as its seed. "The Commodore C64 provided a hardware random number generator, included in its soundchip, the Mos_Technology_SID 6581. Random bytes are fetchable by a read on the correct memory address on the 6581." 

I agree with you regarding the accuracy of my generator and I am concerned about bias. I plan to analyze a large data set to test the generator and perhaps add some filtering and shuffling algorithms to the project. Finally, I am trying to emulate the Global Consciousness Project which uses a hardware generator which is far superior to my model.

Thank you for the input. Any assistance would be greatly appreciated

Best regards...

...Warren

References:

http://www.maths.abdn.ac.uk/~igc/tch/mx3015/notes/node129.html

http://world.std.com/~franl/crypto/random-numbers.html

http://en.wikipedia.org/wiki/Pseudorandom_number_generator

http://en.wikipedia.org/wiki/Hardware_random_number_generator

http://noosphere.princeton.edu/

COMMENT History:

8/21/2007 6:55:18 PM: Cobein

Interesting code, do you have any chance to run statistical tests on it? I was working on a PRNG using hooks, RAW sockets and hardware performance as entropy collectors but I never finished it :D something like Yarrow tho. Well theres a lot of noise in here but no mic to test it, so comments later , 5 *

 
8/21/2007 9:39:06 PM: Multiple Technologies

Its called the egg project, and it has already been done on a much more functional, useful, and efficient manner. This method is really inconvenient as no everyone has a microphone. Why recreate the wheel?

 
8/22/2007 9:36:49 AM: Warren Goff

I've made some adjustments on the audio sampling and randomization. I have done not done any statistical evaluation on it. As for re-inventing the wheel, I wasn't aware of the egg project. I will have to look into it. If it is a Visual Basic project then it will be of utility to me but otherwise, my project allows programmers to utilize this in their projects. Likewise, I am using this in my project to predict the future as in "Moostradamous"

 
8/22/2007 9:42:54 AM: Warren Goff

Thanks Multiple I found what you are talking about and I am humbled. "Global Consciousness Project" http://en.wikipedia.org/wiki/Global_Consciousness_Project [which curiously was edited on Wikopedia by Nostradamus himself ;)]. Still, in my experience, microphones are fairly ubiquitous and the randomization process, although inefficient, is interesting and maybe valuable. Once again, thanks...

 
8/23/2007 1:24:01 AM: Ivan Tellez

If you want a computer do somenting random, you must to tell it how to do it. lol. So, whats the real diference between "True" and "pseudo" Random, both are algorithms, and get a value from a sound captured in a time interval, its in fact, the same thing than get it from a collection of numbers seeded by the internal clock in a specific time. even worse, if you don have a "noisy" ambient, the "Real" random is just a totaly failure. But the internal clock its allways changing its value to seed a random number, even if its considered "pseudo random"

 
