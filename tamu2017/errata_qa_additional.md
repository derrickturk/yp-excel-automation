# TAMU 2017 Workshop: Errata, Unanswered Q&A, Further Reading

## Errata
* Our function to draw random samples from normal distributions (BoundedRandNormal in the example) has a couple of flaws. First, the name: it doesn't apply a bound, so it'd be better named just RandNormal. More importantly, it generates a value from the standard normal distribution (mean of 0, standard deviation of 1) but doesn't apply scaling by the mean and standard distribution it receives as arguments. As a result, we saw 50% of our porosity values (and thus OOIPs) clamped to zero. I've made both corrections to the example posted online.  

## Unanswered Q&A
* The question came up in the workshop of whether the default argument-passing mode in VBA was ByVal (i.e. by-value) or ByRef (i.e. by-reference). It turns out that in VBA the default passing mode for all arguments, regardless of type, is ByRef. As this is rarely what we want, it's another good reason to always be explicit.  

## Further Reading
Several topics were mentioned in the workshop that we didn't have time to discuss in any detail; here are some follow-up resources for more information on these topics.

### Functional programming

*Functional programming* is a style of programming which emphasizes programs as mathematical objects, composed of pure functions. Key ideas include *referential transparency* (the idea that a function call is equivalent to substitution of actual arguments into the function body, with no side effects) and *immutability* (programming with data structures which cannot be changed in-place once they are created); the net effect is to give us the ability of *equational reasoning* about our programs; we don't have to think too hard about "what is *x* at this point in time?" but can rather view our programs more statically and abstractly, as if they were simply mathematical equations (which, in a very real sense, they are).

Another key idea is the use of functions as "first-class" values in a programming system: in functional programming languages, functions may be passed as arguments to functions, returned from functions, and generally manipulated like any other value. This idea is sometimes also referred to from the other end as "higher-order functions": it should be possible to create functions that operate on functions.

While many of these ideas can be used in almost any programming language (Visual Basic included; we tried to make use of pure functions when possible in the workshop), certain languages provide the full suite of features and emphasize a functional programming style. [Haskell](https://www.haskell.org/), [O'Caml](https://ocaml.org/), [F#](http://fsharp.org/), and (to a lesser extent) [Lisp](https://common-lisp.net/) are all popular functional programming languages; they fall along a spectrum from purely-functional Haskell to "as-functional-as-you-like" Lisp.

Functional ideas have achieved some popularity in recent years; many of the "90s scripting languages" (Perl, Python, Ruby, Javascript) allow at least some degree of higher-order functional programming, and even more conservative designs like C++, Java, and C# now provide some functional-programming constructs. (To be quite honest, C++ seems to be on a bender here; it seems to want to become Haskell by 2020 or so.)

You can get a (somewhat muddled) impression of all of this from [Wikipedia](https://en.wikipedia.org/wiki/Functional_programming), but the best way to explore is just to jump in to a functional language, or at least practice following some functional ideas when solving problems in your current programming language.

### The lambda calculus

We explained function application in terms of substitution of actual arguments for formal arguments ("beta reduction"), and used this to justify the "meaninglessness" or "rename-ability" of formal argument names ("alpha substitution"). This justification and terminology comes from the [*lambda calculus*](https://en.wikipedia.org/wiki/Lambda_calculus), a formal system created in the 1930s by mathematician Alonzo Church to study the limits of "computability", in the "a mathematician with enough pencil and paper could find the answer" sense. When digital computers came along a decade later, some enterprising folks noticed that Church's notation for "computable functions" now provided a good foundation for describing, well, *computable* functions.

John McCarthy's LISP language, introduced in 1958 (it lost the ALL CAPS in the 70s or 80s; yes, it's still around), was the first to make the connection really explicit, but many works on formal semantics of programming languages take the lambda calculus as a starting point. It works well as a model for function application in the absence of side effects (see "functional programming", above) and so provides us a clean "textbook" notation with which to reason.

### Compilers and interpreters

We talked briefly about compilers and interpreters in the workshop, as two roughly-defined ends of a spectrum of strategies for "executing programs". The general family of these programs is referred to as "language translators". If you're interested in learning more about how compilers and interpreters work, and why you might want to build one yourself (it's not as hard as it sounds!) you might like the slides for one of my other talks, on [building domain-specific programming languages](https://github.com/derrickturk/dsl-talk).

This talk goes into some detail on a case study where I built a custom programming language for a client to query results from a Monte Carlo simulation, turning it from a "what's the distribution" exercise to a "what's the probability of achieving these specific business goals" tool. We're hoping to publish more on this work as it pertains to the E&P application next year; for now the DSL talk slides give some justification for the approach as well as a sales pitch for the use of functional programming in language translators (and some hints on how to do without). There's also source code in several languages (not VB: [I built it](https://gist.github.com/derrickturk/7ef04a6cfa01d49488735997ac0b1c09), as a potential workshop example, but it's fairly horrible) for a simple integer arithmetic expression interpreter.

### Excel Automation from Other Languages

Excel can be automated from other languages, in much the same way as from VBA. Any language that provides (either inherently or through libraries) bindings to the Microsoft COM protocol (really, an [ABI](https://en.wikipedia.org/wiki/Application_binary_interface)) can use the same Excel object model as VBA.

In the workshop/webinar notes, there's an example in the Day 4 directory which uses Python to re-implement an example from VBA; this shows the general approach. COM bindings are extremely accessible for Python and the various .NET languages. Beyond that, it's theoretically possible to use COM from any language that supports a C-compatible [FFI](https://en.wikipedia.org/wiki/Foreign_function_interface), but building the required infrastructure/boilerplate is quite a bit of work.
