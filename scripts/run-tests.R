lib <- Sys.getenv("R_LIBS_USER")
.libPaths(c(lib, .libPaths()))

testthat::test_local(".", reporter = "summary")
