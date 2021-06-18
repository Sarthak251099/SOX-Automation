import ResetDB
import WindowsACLscan
import WindowsAddUserDetails
import Metrics

def serverScanProd(path):
    env = 'Production'
    ResetDB.resetdb(env)
    metrics = Metrics.TrackWindows(path=path, env=env)
    WindowsACLscan.scan(path, metrics)
    WindowsAddUserDetails.addDetails(path, metrics, env)
    metrics.writeExcel(path=path)
