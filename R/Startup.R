library(devtools)
library(roxygen2)

create_package("C:/Users/TPreijers/Documents/Abstractor", open = F)

use_git()

# use_r("Extractor")
# use_r("Receiver")
# use_r("Builder")

check()
document()


