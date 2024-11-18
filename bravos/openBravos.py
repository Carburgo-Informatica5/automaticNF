import os
import sys

current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)
import infoBravos


class faker:
    def __call__(self, *args, **kwds):
        return self

    def __getattr__(self, name):
        return self


config = {
    "bravos_usr": "caetano.apollo",
    "bravos_pswd": "123",
}
br = infoBravos.bravos(config, m_queue=faker())
br.acquire_bravos(exec="C:\\BravosClient\\BRAVOSClient.exe")
