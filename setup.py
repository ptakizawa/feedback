from setuptools import setup, find_packages

setup(
    name='bd_feedback_ptakizawa',
    version='0.0.1',
    url='https://github.com/mypackage.git',
    author='Peter Takizawa',
    author_email='peter.takizawa@gmail.com',
    description='Summarizes feedback from BlueDogs',
    packages=find_packages(),    
    install_requires=["pandas >=2.2.3", "streamlit >= 1.44.1", "python-docx >= 1.1.2",],
)
