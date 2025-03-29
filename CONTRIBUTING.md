# Contributing to LlamaDocx

Thank you for your interest in contributing to LlamaDocx! This document provides guidelines and instructions for contributing to the project.

## Code of Conduct

This project and everyone participating in it is governed by our Code of Conduct. By participating, you are expected to uphold this code. Please report unacceptable behavior to info@llamasearch.ai.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check the issue list as you might find out that you don't need to create one. When you are creating a bug report, please include as many details as possible:

* Use a clear and descriptive title
* Describe the exact steps which reproduce the problem
* Provide specific examples to demonstrate the steps
* Describe the behavior you observed after following the steps
* Explain which behavior you expected to see instead and why
* Include error messages and stack traces if any
* Include your environment details (OS, Python version, etc.)

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. When creating an enhancement suggestion, please include:

* Use a clear and descriptive title
* Provide a step-by-step description of the suggested enhancement
* Provide specific examples to demonstrate the steps
* Describe the current behavior and explain which behavior you expected to see instead
* Explain why this enhancement would be useful
* List some other tools or applications where this enhancement exists

### Pull Requests

Please follow these steps to have your contribution considered:

1. Follow the [styleguides](#styleguides)
2. After you submit your pull request, verify that all status checks are passing

## Development Process

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Set up your development environment:
   ```bash
   pip install -e .[dev]
   ```
4. Make your changes
5. Run the test suite:
   ```bash
   pytest
   pytest --cov=llamadocx tests/
   ```
6. Run code quality checks:
   ```bash
   black .
   isort .
   flake8
   mypy src/llamadocx
   ```
7. Commit your changes (`git commit -m 'Add amazing feature'`)
8. Push to the branch (`git push origin feature/amazing-feature`)
9. Open a Pull Request

## Styleguides

### Git Commit Messages

* Use the present tense ("Add feature" not "Added feature")
* Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
* Limit the first line to 72 characters or less
* Reference issues and pull requests liberally after the first line

### Python Styleguide

* Follow [PEP 8](https://www.python.org/dev/peps/pep-0008/)
* Use [Black](https://black.readthedocs.io/) for code formatting
* Use [isort](https://pycqa.github.io/isort/) for import sorting
* Use [flake8](https://flake8.pycqa.org/) for linting
* Use [mypy](http://mypy-lang.org/) for type checking
* Write docstrings in Google style

### Documentation Styleguide

* Use [Markdown](https://guides.github.com/features/mastering-markdown/) for documentation
* Reference function and class names using backticks: \`function_name\`
* Include code examples when relevant
* Keep line length to 80 characters or less
* Include docstrings for all public modules, functions, classes, and methods

## Additional Notes

### Issue and Pull Request Labels

* `bug` - Something isn't working
* `enhancement` - New feature or request
* `documentation` - Improvements or additions to documentation
* `good first issue` - Good for newcomers
* `help wanted` - Extra attention is needed
* `invalid` - Something's wrong
* `question` - Further information is requested
* `wontfix` - This will not be worked on

## License

By contributing to LlamaDocx, you agree that your contributions will be licensed under its MIT License. 