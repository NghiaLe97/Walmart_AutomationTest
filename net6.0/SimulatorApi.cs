using SimulationInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static SimulationInterface.ObdSimulator;

namespace SimulatorInterface
{
    /// <summary>
    /// <example>
    /// <br/>step 1: open connection                             :OpenConnection
    /// <br/>step 2: load file simulation                        :Database_LoadFile
    /// <br/>step 3: start device                                :StartDevice
    /// <br/>step 4: keep app running to start simulation job
    /// <br/>step 5: close                                       :Close
    /// </example>
    /// </summary>
    public abstract class SimulatorApi
    {
        public delegate void Callback(string portName, enumOBDMsgType msgtype, string message);
        protected static Callback callbackUpdateText = null;
        /// <summary>
        /// callback function to update console UI when have new event message data comming
        /// </summary>
        /// <param name="cb">callback UI</param>
        public static void setCallbackUpdateConsoleText(Callback cb)
        {
            callbackUpdateText = cb;
        }

        /// <summary>
        /// Open Specific ComportName connection
        /// </summary>
        /// <param name="comportName">ComportName name</param>
        /// <returns>true : success , false : failure</returns>
        public abstract bool OpenConnection(string comportName);

        /// <summary>
        /// close present connection of simulator device and destroy current context obd simulation
        /// </summary>
        public abstract void Close();


        /// <summary>
        /// close present connection of simulator device
        /// </summary>
        public abstract void CloseConnection();

        /// <summary>
        /// Erase current main firmware and jump to bootloader
        /// to enable using upgrade tool to upgrade new simulation fw
        /// </summary>
        public abstract void EraseFW();


        /// <summary>
        /// load specific file simulation to simulator device
        /// return true if success , false: failure
        /// </summary>
        /// <param name="filename">full path file simulation </param>
        /// <param name="updateSetting">after load file simulation will be write setting to device </param>
        /// <returns>true : success , false : failure</returns>
        public abstract bool Database_LoadFile(string filename, bool updateSetting = true);


        /// <summary>
        /// starting simulator device
        /// return true if success , false: failure
        /// </summary>
        public abstract bool StartDevice();

        /// <summary>
        /// stop simulator device
        /// return true if success , false: failure
        /// </summary>
        public abstract bool StopDevice();


        /// <summary>
        /// get present status of simulator device
        /// return true if connected , false: not connected
        /// </summary>
        public abstract bool IsDeviceConnected();

        /// <summary>
        /// get present status of simulator device
        /// return true if started , false: stop
        /// </summary>
        public abstract bool IsDeviceStart();


    }


    public abstract class SimulatorUIApi
    {
        public abstract Dictionary<string,enumprotocol> getProtocolSupport();

        public abstract Dictionary<string, enumDLCPinName> getDLCRx_CanH(enumprotocol protocol);

        public abstract Dictionary<string, enumDLCPinName> getDLCTx_CanL(enumprotocol protocol);

        public abstract void setProtocol(enumprotocol eProtocol);
        public abstract void setProtocol(string Protocol);
        public abstract enumprotocol getProtocol();

        public abstract void setComport(string comportname);
        public abstract string getComport();


        public abstract void setDLCRxCANH(string dlcname, string voltagelevel, bool isInverted, string resistor);
        public abstract structDLCProfile getDLCRxCANH();
        public abstract void setDLCTxCANL(string dlcname, string voltagelevel, bool isInverted, string resistor);
        public abstract structDLCProfile getDLCTxCANL();
    }
}
