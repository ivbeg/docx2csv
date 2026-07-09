# Contributing

Contributions are welcome, and they are greatly appreciated! Every little bit helps.

## Types of Contributions

### Report Bugs

Report bugs at https://github.com/ivbeg/docx2csv/issues.

If you are reporting a bug, please include:

- Your operating system name and version
- Any details about your local setup that may be helpful in troubleshooting
- Detailed steps to reproduce the bug

### Fix Bugs

Look through the GitHub issues for bugs. Anything tagged with "bug" is open to whoever wants to implement it.

### Implement Features

Look through the GitHub issues for features. Anything tagged with "feature" is open to whoever wants to implement it.

### Write Documentation

docx2csv could always use more documentation, whether as part of the official docs, in docstrings, or even on the web in blog posts, articles, and such.

### Submit Feedback

The best way to send feedback is to file an issue at https://github.com/ivbeg/docx2csv/issues.

## Get Started!

1. Fork the `docx2csv` repo on GitHub
2. Clone your fork locally:

    ```bash
    git clone https://github.com/your-username/docx2csv.git
    ```

3. Install your local copy into a virtualenv:

    ```bash
    cd docx2csv
    pip install -e ".[dev]"
    ```

4. Create a branch for local development:

    ```bash
    git checkout -b name-of-your-bugfix-or-feature
    ```

5. Make your changes and add tests
6. Run the tests:

    ```bash
    pytest
    ```

7. Check linting:

    ```bash
    flake8 docx2csv tests
    ```

8. Commit your changes and push your branch:

    ```bash
    git add .
    git commit -m "Your detailed description of your changes"
    git push origin name-of-your-bugfix-or-feature
    ```

9. Submit a pull request through the GitHub website

## Pull Request Guidelines

1. The pull request should include tests
2. If the pull request adds functionality, the docs should be updated
3. Tests should pass for all supported Python versions (3.8-3.12)
