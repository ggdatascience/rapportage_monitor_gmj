\name{bereken_cijfers}
\alias{bereken_cijfers}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{bereken cijfers
%%  ~~function to do ... ~~
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
bereken_cijfers(x)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{x}{
%%     ~~Describe \code{x} here~~
}
}
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
%% Vermeulen, Sander <s.vermeulen@vrln.nl>
%% Michielse, Stijn <Stijn.Michielse@ggdzl.nl>
%% Arne Meeldijk <a.meeldijk@ggdbzo.nl>
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
#Bereken cijfers voor indicator 'Voelt zich gelukkig'
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'SCHOOL')
bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', groepering = 'KLAS')

#Voor meerdere splitsingen is de toegevoegde significantietoets niet geschikt. 
#bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'KLAS')
#bereken_cijfers(data, 'GELUK', omschrijving = 'Voelt zich gelukkig', uitsplitsing = 'GESLACHT', groepering = 'SCHOOL')
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory (show via RShowDoc("KEYWORDS")):
% \keyword{ ~kwd1 }
% \keyword{ ~kwd2 }
% Use only one keyword per line.
% For non-standard keywords, use \concept instead of \keyword:
% \concept{ ~cpt1 }
% \concept{ ~cpt2 }
% Use only one concept per line.
