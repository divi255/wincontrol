import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import os
import yaml

from netaddr import IPNetwork, IPAddress
from contextlib import contextmanager

from flask import Flask, request, Response, jsonify
from functools import wraps

Flask.debug = False

dir_me = os.path.dirname(os.path.realpath(__file__))

app = Flask(__name__)

hosts_allow = []

result_ok = {'ok': True}
result_error = {'ok': False}

config = {}


def load_config():
    config.clear()
    config.update(yaml.load(open(f'{dir_me}/wincontrol.yml').read()))
    if 'access-key' in config:
        config['access-key'] = str(config['access-key'])
    if 'hosts-allow' in config:
        config['hosts-allow'] = [IPNetwork(h) for h in config['hosts-allow']]
    if not isinstance(config.get('command'), dict):
        config['command'] = {}
    try:
        config['host'], config['port'] = config['listen'].split(':', 1)
        config['port'] = int(config['port'])
    except:
        config['host'], config['port'] = '127.0.0.1', 9001


def netacl_match(host, acl):
    if not acl: return True
    for a in acl:
        if IPAddress(host) in a: return True
    return False


def require_auth(f):

    @wraps(f)
    def mf(*args, **kwargs):
        if not netacl_match(request.remote_addr, config.get('hosts-allow')):
            return Response('Access denied\n', 401)
        if request.headers.get('X-Auth-Key') != config.get('access-key'):
            return Response('Invalid key\n', 401)
        return f(*args, **kwargs)

    return mf


@app.route('/state')
@require_auth
def state():
    return jsonify(result_ok)


@app.route('/command/<cmd>', methods=['POST'])
@require_auth
def command(cmd):
    try:
        x = config['command'][cmd]
        if not os.system(x):
            return jsonify(result_ok)
        else:
            return jsonify(result_error)
    except KeyError:
        return Response('Command not defined\n', 404)


# make dummy stdout/stderr, services don't have it and some lib may cause crash
@contextmanager
def no_stdout():
    import sys

    class NoSTDOUT():

        closed = False

        def write(self, text):
            pass

        def flush(self):
            pass

    sys.stdout = NoSTDOUT()
    sys.stderr = NoSTDOUT()


class Service(win32serviceutil.ServiceFramework):
    _svc_name_ = "SimpleWinControlAPI"
    _svc_display_name_ = "Windows Remote Control Simple API"

    def __init__(self, *args):
        win32serviceutil.ServiceFramework.__init__(self, *args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(5)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))

        self.main()

    def main(self):
        try:
            no_stdout()
            load_config()
            app.run(host=config['host'], port=config['port'])
        except:
            import traceback
            open(dir_me + '/wincontrol-error.log',
                 'a').write('{}\n'.format(traceback.format_exc()))


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'app':
        load_config()
        print('Config loaded')
        print('Access key: ' + config['access-key'] if 'access-key' in \
                                 config else 'WARNING: no access key')
        print('Hosts ACL:')
        if 'hosts-allow' in config:
            for h in config['hosts-allow']:
                print(f'  +{h}')
            print('  -', end='')
        else:
            print('  +', end='')
        print('0.0.0.0/0')
        if config.get('command'):
            print('Commands:')
            for k, v in config['command'].items():
                print(f'  POST /command/{k} -> {v}')
        else:
            print('No commands defined')
        print()
        app.run(host=config['host'], port=config['port'])
    else:
        win32serviceutil.HandleCommandLine(Service)
