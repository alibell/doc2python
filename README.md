# doc2python

Ali BELLAMINE - contact@alibellamine.me Last version : 1.0 - 07/02/2021

Main repository : https://gogs.alibellamine.me/alibell/doc2python/

Tool that extract text data from doc file.

## How to install it ?


Clone the current repository :
```
    git clone https://gogs.alibellamine.me/alibell/doc2python
```

Install dependencies with pip.

```
    pip install -r requirements.txt
```

Then install the library :

```
    pip install -e .
```

## How to use it

```
    from doc2python import process

    text = process(path_to_doc)
```