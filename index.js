const { Enumerator, VBArray } = require('JScript')
const { toPosixSep } = require('pathname')
const { get } = require('argv')

const WShell = require('WScript.Shell')
const FSO = require("Scripting.FileSystemObject")
const IWshNetwork2 = require("WScript.Network")

const SWbemLocator = require("WbemScripting.SWbemLocator")
const SWbemServicesEx = SWbemLocator.ConnectServer()

// 定数
const constants = {
    UV_UDP_REUSEADDR: 4,
    dlopen: {},
    errno: {
        E2BIG: 7,
        EACCES: 13,
        EADDRINUSE: 100,
        EADDRNOTAVAIL: 101,
        EAFNOSUPPORT: 102,
        EAGAIN: 11,
        EALREADY: 103,
        EBADF: 9,
        EBADMSG: 104,
        EBUSY: 16,
        ECANCELED: 105,
        ECHILD: 10,
        ECONNABORTED: 106,
        ECONNREFUSED: 107,
        ECONNRESET: 108,
        EDEADLK: 36,
        EDESTADDRREQ: 109,
        EDOM: 33,
        EEXIST: 17,
        EFAULT: 14,
        EFBIG: 27,
        EHOSTUNREACH: 110,
        EIDRM: 111,
        EILSEQ: 42,
        EINPROGRESS: 112,
        EINTR: 4,
        EINVAL: 22,
        EIO: 5,
        EISCONN: 113,
        EISDIR: 21,
        ELOOP: 114,
        EMFILE: 24,
        EMLINK: 31,
        EMSGSIZE: 115,
        ENAMETOOLONG: 38,
        ENETDOWN: 116,
        ENETRESET: 117,
        ENETUNREACH: 118,
        ENFILE: 23,
        ENOBUFS: 119,
        ENODATA: 120,
        ENODEV: 19,
        ENOENT: 2,
        ENOEXEC: 8,
        ENOLCK: 39,
        ENOLINK: 121,
        ENOMEM: 12,
        ENOMSG: 122,
        ENOPROTOOPT: 123,
        ENOSPC: 28,
        ENOSR: 124,
        ENOSTR: 125,
        ENOSYS: 40,
        ENOTCONN: 126,
        ENOTDIR: 20,
        ENOTEMPTY: 41,
        ENOTSOCK: 128,
        ENOTSUP: 129,
        ENOTTY: 25,
        ENXIO: 6,
        EOPNOTSUPP: 130,
        EOVERFLOW: 132,
        EPERM: 1,
        EPIPE: 32,
        EPROTO: 134,
        EPROTONOSUPPORT: 135,
        EPROTOTYPE: 136,
        ERANGE: 34,
        EROFS: 30,
        ESPIPE: 29,
        ESRCH: 3,
        ETIME: 137,
        ETIMEDOUT: 138,
        ETXTBSY: 139,
        EWOULDBLOCK: 140,
        EXDEV: 18,
        WSAEINTR: 10004,
        WSAEBADF: 10009,
        WSAEACCES: 10013,
        WSAEFAULT: 10014,
        WSAEINVAL: 10022,
        WSAEMFILE: 10024,
        WSAEWOULDBLOCK: 10035,
        WSAEINPROGRESS: 10036,
        WSAEALREADY: 10037,
        WSAENOTSOCK: 10038,
        WSAEDESTADDRREQ: 10039,
        WSAEMSGSIZE: 10040,
        WSAEPROTOTYPE: 10041,
        WSAENOPROTOOPT: 10042,
        WSAEPROTONOSUPPORT: 10043,
        WSAESOCKTNOSUPPORT: 10044,
        WSAEOPNOTSUPP: 10045,
        WSAEPFNOSUPPORT: 10046,
        WSAEAFNOSUPPORT: 10047,
        WSAEADDRINUSE: 10048,
        WSAEADDRNOTAVAIL: 10049,
        WSAENETDOWN: 10050,
        WSAENETUNREACH: 10051,
        WSAENETRESET: 10052,
        WSAECONNABORTED: 10053,
        WSAECONNRESET: 10054,
        WSAENOBUFS: 10055,
        WSAEISCONN: 10056,
        WSAENOTCONN: 10057,
        WSAESHUTDOWN: 10058,
        WSAETOOMANYREFS: 10059,
        WSAETIMEDOUT: 10060,
        WSAECONNREFUSED: 10061,
        WSAELOOP: 10062,
        WSAENAMETOOLONG: 10063,
        WSAEHOSTDOWN: 10064,
        WSAEHOSTUNREACH: 10065,
        WSAENOTEMPTY: 10066,
        WSAEPROCLIM: 10067,
        WSAEUSERS: 10068,
        WSAEDQUOT: 10069,
        WSAESTALE: 10070,
        WSAEREMOTE: 10071,
        WSASYSNOTREADY: 10091,
        WSAVERNOTSUPPORTED: 10092,
        WSANOTINITIALISED: 10093,
        WSAEDISCON: 10101,
        WSAENOMORE: 10102,
        WSAECANCELLED: 10103,
        WSAEINVALIDPROCTABLE: 10104,
        WSAEINVALIDPROVIDER: 10105,
        WSAEPROVIDERFAILEDINIT: 10106,
        WSASYSCALLFAILURE: 10107,
        WSASERVICE_NOT_FOUND: 10108,
        WSATYPE_NOT_FOUND: 10109,
        WSA_E_NO_MORE: 10110,
        WSA_E_CANCELLED: 10111,
        WSAEREFUSED: 10112
    },
    signals: {
        SIGHUP: 1,
        SIGINT: 2,
        SIGILL: 4,
        SIGABRT: 22,
        SIGFPE: 8,
        SIGKILL: 9,
        SIGSEGV: 11,
        SIGTERM: 15,
        SIGBREAK: 21,
        SIGWINCH: 28
    },
    priority: {
        PRIORITY_LOW: 19,
        PRIORITY_BELOW_NORMAL: 10,
        PRIORITY_NORMAL: 0,
        PRIORITY_ABOVE_NORMAL: -7,
        PRIORITY_HIGH: -14,
        PRIORITY_HIGHEST: -20
    }
}
const EOL = "\r\n"
const devNull = "\\\\.\\nul"

/**
 * Node.js バイナリがコンパイルされたオペレーティング システムの CPU アーキテクチャを返します。 可能な値は、"arm", "arm64", "ia32", "mips", "mipsel", "ppc", "ppc64", "s390", "s390x", および "x64"です。
 * @returns {string}
 */
function arch() {
    if (get('arch') === 'AMD64') return 'x64'
    return WShell.ExpandEnvironmentStrings('%PROCESSOR_ARCHITECTURE%')
}

/**
 * プログラムが使用するデフォルトの並列処理量の推定値を返します。 常にゼロより大きい値を返します。
 * @returns {number} cpuのプロセッサー数
 */
function availableParallelism() {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("SELECT * FROM Win32_Processor")
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)
    const cpu = SWbemObjectSetEx[0]
    return cpu.NumberOfLogicalProcessors
}

/**
 * 各論理 CPU コアに関する情報を含むオブジェクトの配列を返します。 /proc ファイル システムが利用できない場合など、CPU 情報が利用できない場合、配列は空になります。
 * @returns {cpu[]} cpu
 */
function cpus() {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("SELECT * FROM Win32_Processor")
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)
    const cpu = SWbemObjectSetEx[0]

    // cpu情報を配列で受け取れないため、プロセッサーの数だけダミーを生成
    return new Array(cpu.NumberOfLogicalProcessors)
        .fill({
            model: cpu.Name,
            speed: cpu.MaxClockSpeed,
            times: {
                user: 0,
                nice: 0,
                sys: 0,
                idle: 0,
                irq: 0
            }
        })
    return ''
}

/**
 * 空きシステム メモリの量をバイト単位で整数として返します
 * @returns {number} 利用可能物理メモリ
 */
function freemem() {
    // 空きシステム メモリの量をバイト単位で整数として返します。
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * FROM Win32_OperatingSystem")
    const SWbemObjectEx = new Enumerator(SWbemObjectSet)[0]

    /*
    console.log("合計物理メモリ")
    console.log(() => SWbemObjectEx.TotalVisibleMemorySize)

    console.log("合計仮想メモリ")
    console.log(() => SWbemObjectEx.TotalVirtualMemorySize)

    console.log("利用可能物理メモリ")
    console.log(() => SWbemObjectEx.FreePhysicalMemory)

    console.log("利用可能仮想メモリ")
    console.log(() => SWbemObjectEx.FreeVirtualMemory)

    console.log("コミットチャージ合計")
    console.log(() => SWbemObjectEx.SizeStoredInPagingFiles - SWbemObjectEx.FreeSpaceInPagingFiles)

    console.log("コミットチャージ制限値")
    console.log(() => SWbemObjectEx.SizeStoredInPagingFiles)

    console.log("ページングファイルにマップできるサイズ")
    console.log(() => SWbemObjectEx.FreeSpaceInPagingFiles)

    console.log("仮説 合計物理メモリ - 利用可能物理メモリ")
    console.log(() => SWbemObjectEx.TotalVisibleMemorySize - SWbemObjectEx.FreePhysicalMemory)
    */
    return SWbemObjectEx.FreePhysicalMemory
}

/**
 * pid で指定されたプロセスのスケジューリング優先度を返します。 pid が指定されていないか、0 の場合は、現在のプロセスの優先度が返されます。
 * @param {number} pid - プロセスID
 * @returns {number} 優先度
 */
function getPriority(pid) {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * FROM Win32_Process")
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)

    // 現在のプロセスIDを取得
    if (pid == null || pid === 0) pid = SWbemServicesEx.Get(`Win32_Process.Handle='${WShell.Exec("mshta.exe").ProcessID}'`).ParentProcessId

    let result = 0
    SWbemObjectSetEx
        .filter(SWbemObjectEx => SWbemObjectEx.ProcessId === pid)
        .forEach(SWbemObjectEx => result = SWbemObjectEx.Priority)

    return result
}

/**
 * オペレーティング システムのホスト名を文字列として返します。
 * @returns {string} ホスト名
 */
function hostname() {
    return IWshNetwork2.ComputerName.toLowerCase()
}

/**
 * 1、5、15 分間の負荷平均を含む配列を返します。(未実装)
 * @returns {number[]} 整数値の配列
 */
function loadavg() {
    return [0, 0, 0]
}

/**
 * ネットワーク アドレスが割り当てられたネットワークインターフェイスを含むオブジェクトを返します。返されたオブジェクトの各キーはネットワークインターフェイスを識別します。 関連付けられた値は、割り当てられたネットワークアドレスをそれぞれ記述するオブジェクトの配列です。(実装不備あり)
 * @returns {object[]} ネットワークインターフェイスの配列
 */
function networkInterfaces() {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * From Win32_NetworkAdapterConfiguration" /*Where IPEnabled = True"*/)
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)

    const result = {}
    SWbemObjectSetEx.forEach(SWbemObjectEx => {
        const caption = SWbemObjectEx.Caption
        let res = [{}]

        if (SWbemObjectEx.IPSubnet != null) {
            const IPSubnet = new VBArray(SWbemObjectEx.IPSubnet)
            IPSubnet.forEach((netmask, i) => {
                res[i] = res[i] || {}
                res[i].netmask = netmask
            })
        } else res[0].netmask = ""

        if (SWbemObjectEx.IPAddress != null) {
            const IPAddress = new VBArray(SWbemObjectEx.IPAddress)
            IPAddress.forEach((address, i) => {
                res[i] = res[i] || {}
                res[i].address = address
                res[i].family = address.includes('::') ? 'IPv6' : 'IPv4'
            })
        } else {
            res[0].address = ""
            res[0].family = ""
        }

        const mac = SWbemObjectEx.MACAddress
        res = res.map(item => {
            item.mac = mac
            return item
        })

        result[caption] = res
    })

    return result
}

/**
 * pid で指定されたプロセスのスケジューリング優先度の設定を試みます。 pid が指定されていないか、0 の場合は、現在のプロセスのプロセス ID が使用されます。優先度入力は、-20 (高優先度) から 19 (低優先度) までの整数である必要があります。 Unix の優先度レベルと Windows の優先度クラスの違いにより、優先度は os.constants.priority の 6 つの優先度定数のいずれかにマップされます。 プロセスの優先度レベルを取得する場合、この範囲マッピングにより Windows では戻り値が若干異なる場合があります。 混乱を避けるために、優先度を優先度定数のいずれかに設定します。Windows では、優先順位を PRIORITY_HIGHEST に設定するには、昇格されたユーザー権限が必要です。 それ以外の場合、設定された優先度は黙って PRIORITY_HIGH に減らされます。(要テスト)
 * @param {number} pid - プロセスID
 * @param {number} priority - 優先度
 */
function setPriority(pid = getPriority(), priority = constants.priority.PRIORITY_NORMAL) {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * FROM Win32_Process")
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)

    SWbemObjectSetEx.forEach(SWbemObjectEx => {
        if (SWbemObjectEx.ProcessId === pid) SWbemObjectEx.SetPriority(priority)
    })
}

/**
 * オペレーティング システムの一時ファイルのデフォルトディレクトリを文字列として返します。
 * @returns {string} ディレクトリパス
 */
function tmpdir() {
    return toPosixSep(FSO.GetSpecialFolder(2).Path)
}

/**
 * システムメモリの合計量をバイト単位で整数として返します。
 * @returns {number} 合計物理メモリ
 */
function totalmem() {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * FROM Win32_OperatingSystem")
    const SWbemObjectEx = new Enumerator(SWbemObjectSet)[0]

    return SWbemObjectEx.TotalVisibleMemorySize * 1024
}

/**
 * uname(3) によって返されたオペレーティングシステム名を返します。 "Windows_NT" を返します。
 * @returns {string} オペレーティングシステム名
 */
function type() {
    return 'Windows_NT'
}

/**
 * 現在有効なユーザーに関する情報を返します。 POSIX プラットフォームでは、これは通常、パスワード ファイルのサブセットです。 返されるオブジェクトには、ユーザー名、uid、gid、shell、homedir が含まれます。 Windows では、uid フィールドと gid フィールドは -1 で、shell は null です。
 * @param {object} 出力の指定エンコーディング(未実装)
 * @returns {object} ユーザー情報
 */
function userInfo(options = { encoding: 'UTF-8' }) {
    const info = {
        uid: -1,
        gid: -1,
        username: WShell.ExpandEnvironmentStrings('%USERNAME%'),
        homedir: toPosixSep(WShell.ExpandEnvironmentStrings('%HOMEDRIVE%') + WShell.ExpandEnvironmentStrings('%HOMEPATH%')),
        shell: null
    }
    return info
}

/**
 * システム稼働時間を秒数で返します。
 * @returns {number} システム稼働時間
 */
function uptime() {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * From Win32_PerfFormattedData_PerfOS_System")
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)

    return SWbemObjectSetEx.reduce((acc, curr) => {
        if (curr.SystemUptime) return curr.SystemUptime
    }, '')
}

/**
 * カーネルのバージョンを識別する文字列を返します。
 * @returns {string} カーネルのバージョン
 */
function version() {
    const SWbemObjectSet = SWbemServicesEx.ExecQuery("Select * From Win32_OperatingSystem")
    const SWbemObjectSetEx = new Enumerator(SWbemObjectSet)

    return SWbemObjectSetEx.reduce((acc, curr) => {
        if (curr.Caption) return curr.Caption
    }, '')
}

/**
 * Windows では RtlGetVersion() が使用され、それが使用できない場合は GetVersionExW() が使用されます。(未実装)
 * @returns {string} マシンタイプ
 */
function machine() {
    return "x86_64"
}

module.exports = {
    arch,
    availableParallelism,
    cpus,
    freemem,
    getPriority,
    hostname,
    loadavg,
    networkInterfaces,
    setPriority,
    tmpdir,
    totalmem,
    type,
    userInfo,
    uptime,
    version,
    machine,
    constants,
    EOL,
    devNull
}