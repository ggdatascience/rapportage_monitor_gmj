\name{content_genereren}
\alias{content_genereren}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{content genereren
%%  ~~function to do ... ~~
}
\description{
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
content_genereren(data, omschrijving, indicator, groepering, type)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{data}{
%%     ~~Describe \code{data} here~~
}
  \item{omschrijving}{
%%     ~~Describe \code{omschrijving} here~~
}
  \item{indicator}{
%%     ~~Describe \code{indicator} here~~
}
  \item{groepering}{
%%     ~~Describe \code{groepering} here~~
}
  \item{type}{
%%     ~~Describe \code{type} here~~
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
%%  ~~who you are~~
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
##---- Should be DIRECTLY executable !! ----
##-- ==>  Define data, use random,
##--	or do  help(data=index)  for the standard data sets.

## The function is currently defined as
function (data, omschrijving, indicator, groepering, type)
{
    if (type == "tekst") {
        omschrijving
    }
    else if (type == "percentage") {
        mean(data[[indicator]] * 100, na.rm = T) \%>\% round(0) \%>\%
            paste0("\%")
    }
    else if (type == "staafgrafiek") {
        ms_barchart(data = data, x = groepering, y = indicator)
    }
  }
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
