from setuptools import setup, find_packages


setup(name='doc2python', 
      version='0.0.1',
      license='',
      author='Ali BELLAMINE',
      author_email='contact@alibellamine.me',
      description='Extract text from doc file.',
      long_description=open('README.md').read(),
      include_package_data=True,
      packages = find_packages(include=["doc2python"]),
      package_data = {
          "":["data/*"]
      }
)