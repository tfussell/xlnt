# Contributing to xlnt

xlnt welcomes contributions from everyone regardless of skill level (provided you can write C++ or documentation).

## Getting Started

Look through the list of issues to find something interesting to work on. Help is appreciated with any issues, but important timely issues are labeled as "help wanted". Issues labeled "docs" might be good for those who want to contribute without having to know too much C++. You might also find something that the code is missing without an associated issue. That's fine to work on to, but it might be best to make an issue first in case someone else is working on it.

## Contributions

Contributions to xlnt should be made in the form of pull requests on GitHub. Each pull request will be reviewed and either merged into the current development branch or given feedback for changes that would be required to do so. 

All code in this repository is under the MIT License. You should agree to these terms before submitting any code to xlnt.

## Pull Request Checklist

- Branch from the head of the current development branch. Until version 1.0 is released, this the master branch.

- Commits should be as small as possible, while ensuring that each commit is correct independently (i.e. each commit should compile and pass all tests). Commits that don't follow the coding style indicated in .clang-format (e.g. indentation) are less likely to be accepted until they are fixed.

- If your pull request is not getting reviewed or you need a specific person to review it, you can @-reply a reviewer asking for a review in the pull request or a comment.

- Add tests relevant to the fixed defect or new feature. It's best to do this before making any changes, make sure that the tests fail, then make changes ensuring that it ultimately passes the tests (i.e. TDD). xlnt uses cxxtest for testing. Tests are contained in a tests directory inside each module (e.g. source/workbook/tests/test_workbook.hpp) in the form of a header file. Each test is a separate function with a name that starts like "test_". See http://cxxtest.com/guide.html for information about CxxTest or take a look at existing tests.

## Conduct

Just try to be nice--we're all volunteers here.

## Communication

Add a comment to an existing issue on GitHub, open a new issue for defects or feature requests, or contact @tfussell if you want.
