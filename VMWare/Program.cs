using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using AppUtil;
using Vim25Api;

namespace VMWare
{
    public class Program
    {
        public static void Main(string[] args)
        {

            if (File.Exists(@"C:\Temp\Vmlog.log"))
            {
                System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\nVMList - args: ");
                args.ToList().ForEach(item => System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", item));
            }


            //string result = "";
            StringWriter sw = new StringWriter();
            DataTable dt = new DataTable("resultSet");
            dt.Columns.Add("Result", typeof(String));

            try
            {
                if (args.Length == 0)
                {

                    throw new Exception("Incorrect parameter count");
                }
                else
                {

                    string arg = String.Join(" ", args.Select(p => p.ToString()).ToArray());
                    DataTable dtParams = new DataTable("");
                    StringReader reader = new StringReader(arg);
                    dtParams.ReadXml(reader);
                    DataRow params_row = dtParams.Rows[0];
                    Dictionary<string, string> dicHostMap = new Dictionary<string, string>();

                    AppUtil.AppUtil cb = null;

                    try
                    {
                        string[] hostMap = System.Configuration.ConfigurationManager.AppSettings["HostMap"].Split(';');
                        dicHostMap.Add(hostMap[0], hostMap[1]);
                    }
                    catch (Exception e) { } //failed to load the config file


                    string[] appOptions = new string[8];
                    appOptions[0] = "--username";
                    appOptions[1] = params_row["UserName"].ToString();

                    if (appOptions[1].Contains("\\"))
                    {
                        appOptions[1] = appOptions[1].Substring(appOptions[1].LastIndexOf("\\") + 1);
                    }


                    appOptions[2] = "--password";
                    appOptions[3] = params_row["Password"].ToString();
                    appOptions[4] = "--url";
                    if (dicHostMap.Keys.Contains(params_row["HostName"].ToString()))
                    {
                        appOptions[5] = "https://" + dicHostMap[params_row["HostName"].ToString()] + "/sdk";
                    }
                    else
                    {
                        appOptions[5] = "https://" + params_row["HostName"].ToString() + "/sdk";
                    }
                    appOptions[6] = "--ignorecert";
                    appOptions[7] = "--disablesso";

                    try
                    {
                        cb = AppUtil.AppUtil.initialize("VMCommands", null, appOptions);
                        cb.connect();
                    }
                    catch (Exception e)
                    {
                        if (e.Message.StartsWith("Cannot complete login due to an incorrect user name or password"))
                        {
                            throw;
                        }
                        else
                        {
                            throw new Exception("Could not connect to host", e);
                        }
                    }

                  
                    switch (params_row["Command"].ToString())
                    {
                        //commands
                        #region Simple commands
                        case "VMList":
                            {

                                if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nVMList - Starting VM List"); }
                                DataTable dtVms = new DataTable("resultSet");
                                dtVms.Columns.Add("Virtual machine");
                                dtVms.Columns.Add("Uuid");
                                dtVms.Columns.Add("Host");
                                dtVms.Columns.Add("Folder");

                                ArrayList hosts = cb.getServiceUtil().GetDecendentMoRefs(null, "HostSystem");
                                if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nList - Got hosts"); }
                                foreach (object o_host in hosts)
                                {
                                    ManagedObjectReference host = (ManagedObjectReference)((object[])o_host)[0];
                                    string sHostName = (string)cb.getServiceUtil().GetDynamicProperty(host, "name");

                                    if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nList - Checking VMs for " + host.Value); }
                                    System.Collections.ArrayList vms = cb.getServiceUtil().GetDecendentMoRefs(host, "VirtualMachine");
                                    if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nList - Found " + vms.Count + "VMs"); }
                                    foreach (object[] objRef in vms)
                                    {
                                        DataRow r = dtVms.NewRow();

                                        r["Virtual machine"] = objRef[1].ToString();
                                        if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nList - Getting uuid for VM " + objRef[1].ToString()); }
                                        ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", objRef[1].ToString());
                                        if (cb.getServiceUtil().GetDynamicProperty(vm, "config.uuid") != null)
                                        {
                                            r["Uuid"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.uuid").ToString();
                                        }
                                       
                                        r["Folder"] = GetFolder(ref cb, vm);

                                        r["Host"] = sHostName;
                                        if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nList - Getting uuid for VM " + objRef[1].ToString() + " Done"); }
                                        dtVms.Rows.Add(r);
                                    }
                                }

                                //Sort
                                if (dtVms.Rows.Count > 0)
                                {
                                    if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nList - Sorting"); }
                                    DataView dv = dtVms.DefaultView;
                                    dv.Sort = "Virtual machine";
                                    dtVms = dv.ToTable();
                                }
                                dt = dtVms;
                                break;
                            }
                        case "VMHostList":
                            {
                                if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nVMHostList - Starting VM Host List"); }

                                System.Collections.ArrayList hosts = cb.getServiceUtil().GetDecendentMoRefs(null, "HostSystem");

                                if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nVMHostList - Data received"); }

                                DataTable dtVms = new DataTable("resultSet");
                                dtVms.Columns.Add("Host");
                                //dtVms.Columns.Add("uuid");

                                foreach (Object[] objRef in hosts)
                                {
                                    DataRow r = dtVms.NewRow();
                                    r["Host"] = objRef[1].ToString();
                                   
                                    dtVms.Rows.Add(r);
                                }

                                //Sort
                                if (dtVms.Rows.Count > 0)
                                {
                                    DataView dv = dtVms.DefaultView;
                                    dv.Sort = "Host";
                                    dtVms = dv.ToTable();
                                }
                                dt = dtVms;

                                if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\nVMHostList - result build"); }

                                break;
                            }
                        case "VMExists":
                            {
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    dt.Rows.Add(dt.NewRow()["Result"] = "True");
                                }
                                else
                                {
                                    dt.Rows.Add(dt.NewRow()["Result"] = "False");
                                }
                                break;
                            }
                        case "VMInfo":
                            {
                               
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    DataTable dtres = new DataTable("resultSet");
                                    dtres.Columns.Add("VM name", typeof(String));
                                    dtres.Columns.Add("VM uuid", typeof(String));
                                    dtres.Columns.Add("Guest OS", typeof(String));
                                    dtres.Columns.Add("Guest state", typeof(String));
                                    dtres.Columns.Add("Guest tools status", typeof(String));
                                    dtres.Columns.Add("Guest ip", typeof(String));
                                    dtres.Columns.Add("VM number of CPUs", typeof(String));
                                    dtres.Columns.Add("VM number of cores per socket", typeof(String));
                                    dtres.Columns.Add("VM memory", typeof(String));
                                    dtres.Columns.Add("Host name", typeof(String));
                                    dtres.Columns.Add("IsTemplate", typeof(String));

                                    DataRow r = dtres.NewRow();
                                    r["VM name"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.name").ToString();
                                    r["VM uuid"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.uuid").ToString();
                                    r["Guest OS"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.guestFullName").ToString();
                                    r["Guest state"] = cb.getServiceUtil().GetDynamicProperty(vm, "guest.guestState").ToString();
                                    r["Guest tools status"] = cb.getServiceUtil().GetDynamicProperty(vm, "guest.toolsRunningStatus").ToString();

                                    ManagedObjectReference host = GetHostForVM(vm, ref cb);
                                    if (host != null)
                                    {
                                        r["Host name"] = cb.getServiceUtil().GetDynamicProperty(host, "name");
                                    }

                                    object ip_property = cb.getServiceUtil().GetDynamicProperty(vm, "guest.ipAddress");
                                    if (ip_property != null)
                                    { r["Guest ip"] = ip_property.ToString(); }
                                    else
                                    { r["Guest ip"] = ""; }

                                    r["IsTemplate"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.template").ToString();

                                    r["VM number of CPUs"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.hardware.numCPU").ToString();
                                    r["VM number of cores per socket"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.hardware.numCoresPerSocket").ToString();
                                    r["VM memory"] = cb.getServiceUtil().GetDynamicProperty(vm, "config.hardware.memoryMB").ToString();
                                    dtres.Rows.Add(r);

                                    dt = dtres;
                                   
                                }
                                else
                                {
                                   
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMRename":
                            {
                               
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection()._service.Rename_Task(vm, params_row["NewVMName"].ToString());
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                  
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                       
                        case "VMModifyCpu":
                            {

                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                //get min max cpu count
                                if (vm == null)
                                {
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                ManagedObjectReference hostmor = GetHostForVM(vm, ref cb);
                                ManagedObjectReference crmor = cb.getServiceUtil().GetMoRefProp(hostmor, "parent");
                                ComputeResourceSummary rps = (ComputeResourceSummary)cb.getServiceUtil().GetDynamicProperty(crmor, "summary");
                                int newCpuNum = -1;

                                if (!params_row.Table.Columns.Contains("CpuCount")
                                    || !Int32.TryParse(params_row["CpuCount"].ToString(), out newCpuNum))
                                {
                                    throw new Exception("Invalid number of CPUs");
                                }
                                if (rps == null
                                    || rps.numCpuThreads < newCpuNum
                                    || newCpuNum < 1
                                    )
                                { throw new Exception("Illegal number of CPUs :" + newCpuNum.ToString()); }


                                VirtualMachineConfigSpec vmConfigSpec = new VirtualMachineConfigSpec();
                                vmConfigSpec.numCPUs = newCpuNum;
                                vmConfigSpec.numCPUsSpecified = true;
                                ManagedObjectReference taskmor = cb._connection._service.ReconfigVM_Task(vm, vmConfigSpec);
                                dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));

                                break;
                            }
                        case "VMModifyMemory":
                            {
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm == null)
                                {
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }

                                int newMemoryAmount = -1;

                                if (!params_row.Table.Columns.Contains("NewMemorySizeMB")
                                     || !Int32.TryParse(params_row["NewMemorySizeMB"].ToString(), out newMemoryAmount))
                                {
                                    throw new Exception("Invalid memory amount");
                                }
                                if (newMemoryAmount < 1)
                                { throw new Exception("Illegal memory amount:" + newMemoryAmount.ToString()); }

                                VirtualMachineConfigSpec vmConfigSpec = new VirtualMachineConfigSpec();

                                vmConfigSpec.memoryMB = newMemoryAmount;
                                vmConfigSpec.memoryMBSpecified = true;

                                ManagedObjectReference taskmor = cb._connection._service.ReconfigVM_Task(vm, vmConfigSpec);
                                dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));

                                break;
                            }
                        case "VMMountCD":
                            {
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm == null)
                                {
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }

                                VirtualMachineConfigSpec vmConfigSpec = new VirtualMachineConfigSpec();
                                VirtualDeviceConfigSpec cdSpec = new VirtualDeviceConfigSpec();


                                //String ops = cb.get_option("operation");
                                VirtualMachineConfigInfo vmConfigInfo = (VirtualMachineConfigInfo)cb.getServiceUtil().GetDynamicProperty(vm, "config");

                                if (params_row["Operation"].ToString() == "add")
                                {
                                    ManagedObjectReference envBrowser = (ManagedObjectReference)cb.getServiceUtil().GetDynamicProperty(vm, "environmentBrowser");
                                    cdSpec.operation = VirtualDeviceConfigSpecOperation.add;
                                    cdSpec.operationSpecified = true;
                                    DatastoreSummary dsum = getDataStoreSummary(vm, ref cb);


                                    VirtualCdrom cdrom = new VirtualCdrom();

                                    VirtualCdromIsoBackingInfo cdDeviceBacking = new VirtualCdromIsoBackingInfo();
                                    cdDeviceBacking.datastore = dsum.datastore;
                                    cdDeviceBacking.fileName = "[" + dsum.name + "] " + params_row["ISOPath"].ToString();

                                    VirtualDevice[] defaultDevs = null;
                                    VirtualMachineConfigOption cfgOpt = cb.getConnection()._service.QueryConfigOption(envBrowser, null, null);
                                    if (cfgOpt == null)
                                    {
                                        throw new Exception("No VirtualHardwareInfo found in ComputeResource");
                                    }
                                    defaultDevs = cfgOpt.defaultDevice;
                                    if (defaultDevs == null)
                                    {
                                        throw new Exception("No Datastore found in ComputeResource");
                                    }
                                    VirtualDevice ideCtlr = null;
                                    for (int di = 0; di < defaultDevs.Length; di++)
                                    {
                                        if (defaultDevs[di].GetType().Name.Equals("VirtualIDEController"))
                                        {
                                            ideCtlr = defaultDevs[di];
                                            break;
                                        }
                                    }

                                    cdrom.backing = cdDeviceBacking;
                                    cdrom.controllerKey = ideCtlr.key;
                                    cdrom.controllerKeySpecified = true;
                                    cdrom.unitNumber = -1;
                                    cdrom.unitNumberSpecified = true;
                                    cdrom.key = -100;

                                    cdSpec.device = cdrom;
                                }
                                
                                if (cdSpec != null)
                                {
                                    VirtualDeviceConfigSpec[] cdSpecArray = { cdSpec };
                                    vmConfigSpec.deviceChange = cdSpecArray;
                                }
                                ManagedObjectReference taskmor = cb._connection._service.ReconfigVM_Task(vm, vmConfigSpec);
                                dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));

                                break;
                            }
                        case "VMAddDisk":
                            {
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm == null)
                                {
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }

                                VirtualMachineConfigSpec vmConfigSpec = new VirtualMachineConfigSpec();
                                VirtualDeviceConfigSpec vdiskSpec = getDiskDeviceConfigSpec(ref cb, vm);
                                if (vdiskSpec != null)
                                {
                                    VirtualMachineConfigInfo vmConfigInfo = (VirtualMachineConfigInfo)cb.getServiceUtil().GetDynamicProperty(vm, "config");
                                    int ckey = -1;
                                    VirtualDevice[] test = vmConfigInfo.hardware.device;
                                    for (int k = 0; k < test.Length; k++)
                                    {
                                        if (test[k].deviceInfo.label.Equals(
                                           "SCSI Controller 0"))
                                        {
                                            ckey = test[k].key;
                                        }
                                    }

                                    if (ckey == -1)
                                    {
                                        int diskCtlrKey = 1;
                                        VirtualDeviceConfigSpec scsiCtrlSpec = new VirtualDeviceConfigSpec();
                                        scsiCtrlSpec.operation = VirtualDeviceConfigSpecOperation.add;
                                        scsiCtrlSpec.operationSpecified = true;
                                        VirtualLsiLogicController scsiCtrl = new VirtualLsiLogicController();
                                        scsiCtrl.busNumber = 0;
                                        scsiCtrlSpec.device = scsiCtrl;
                                        scsiCtrl.key = diskCtlrKey;
                                        scsiCtrl.sharedBus = VirtualSCSISharing.physicalSharing;
                                        String ctlrType = scsiCtrl.GetType().Name;
                                        vdiskSpec.device.controllerKey = scsiCtrl.key;
                                        VirtualDeviceConfigSpec[] vdiskSpecArray = { scsiCtrlSpec, vdiskSpec };
                                        vmConfigSpec.deviceChange = vdiskSpecArray;
                                    }
                                    else
                                    {
                                        vdiskSpec.device.controllerKey = ckey;
                                        VirtualDeviceConfigSpec[] vdiskSpecArray = { vdiskSpec };
                                        vmConfigSpec.deviceChange = vdiskSpecArray;
                                    }
                                }
                                else
                                {
                                    return;
                                }

                                break;
                            }
                        case "VMResizeDisk":
                            {
                                //System.Diagnostics.Debugger.Launch();

                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm == null)
                                {
                                    throw new Exception(params_row["VMName"].ToString() + " doesn't exist");
                                }

                                VirtualMachineConfigSpec vmConfigSpec = new VirtualMachineConfigSpec();
                                VirtualDeviceConfigSpec diskSpec = null;
                                VirtualDisk disk = null;

                                VirtualMachineConfigInfo vmConfigInfo = (VirtualMachineConfigInfo)cb.getServiceUtil().GetDynamicProperty(vm, "config");
                                VirtualDiskFlatVer2BackingInfo diskfileBacking = new VirtualDiskFlatVer2BackingInfo();

                                VirtualHardware vh = vmConfigInfo.hardware;
                                VirtualDevice[] test = vh.device;

                                int size = -1;
                                if (!Int32.TryParse(params_row["DiskSize"].ToString(), out size))
                                { throw new Exception("Invalid disk size"); }

                                for (int k = 0; k < test.Length; k++)
                                {
                                    Description desc = test[k].deviceInfo;
                                    if (string.Compare(desc.label, params_row["DiskName"].ToString(), true) == 0)
                                    {
                                        disk = (VirtualDisk)test[k];
                                    }
                                }

                                if (disk != null)
                                {
                                    diskSpec = new VirtualDeviceConfigSpec();
                                    disk.capacityInKB = size;
                                    diskSpec.operation = VirtualDeviceConfigSpecOperation.edit;
                                    diskSpec.operationSpecified = true;
                                    diskSpec.device = disk;
                                }
                                else
                                { throw new Exception("Disk " + params_row["DiskName"].ToString() + " doesn't exist"); }

                                if (diskSpec != null)
                                {
                                    VirtualDeviceConfigSpec[] vdiskSpecArray = { diskSpec };
                                    vmConfigSpec.deviceChange = vdiskSpecArray;
                                }

                                ManagedObjectReference tmor = cb._connection._service.ReconfigVM_Task(vm, vmConfigSpec);
                                dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, tmor));

                                break;
                            }
                        case "VMDeleteConfigurationSpec":
                            {
                                if (cb._connection._sic.customizationSpecManager == null)
                                { throw new Exception("Missing customization spec manager"); }

                                cb._connection._service.DeleteCustomizationSpec(cb._connection._sic.customizationSpecManager, params_row["SpecName"].ToString());
                                dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                break;
                            }
                        case "VMCreateConfigurationSpec":
                            {
                                //System.Diagnostics.Debugger.Launch();
                                if (cb._connection._sic.customizationSpecManager == null)
                                { throw new Exception("Missing customization spec manager"); }

                                if (!params_row.Table.Columns.Contains("SpecName") || params_row["SpecName"] == null)
                                { throw new Exception("Missing Config Spec name"); }

                                bool exists = cb._connection._service.DoesCustomizationSpecExist(cb._connection._sic.customizationSpecManager,
                                                                                                    params_row["SpecName"].ToString());
                                //CustomizationSpecItem it = cb._connection._service.XmlToCustomizationSpecItem(
                                //                                                    cb._connection._sic.customizationSpecManager, xml);

                                CustomizationSpecItem it = new CustomizationSpecItem();
                                CustomizationSpecInfo info = new CustomizationSpecInfo();
                                it.spec = new CustomizationSpec();

                                object o = cb.getServiceUtil().GetDynamicProperty(cb._connection._sic.customizationSpecManager, "encryptionKey");
                                sbyte[] bytes = (sbyte[])o;
                                it.spec.encryptionKey = bytes;
                                it.spec.globalIPSettings = new CustomizationGlobalIPSettings();

                                //Name
                                info.name = params_row["SpecName"].ToString();

                                //Type
                                if (params_row.Table.Columns.Contains("ConfigSpecType"))
                                { info.type = params_row["ConfigSpecType"].ToString(); }
                                else { info.type = "Windows"; }

                                info.description = "";
                                it.info = info;
                                //LastUpdateTime
                                //Description
                                //changeVersion

                                //it.spec.identity
                                CustomizationSysprep id = new CustomizationSysprep();
                                CustomizationUserData userData = new CustomizationUserData();
                                if (params_row.Table.Columns.Contains("PersonName"))
                                { userData.fullName = params_row["PersonName"].ToString(); }

                                if (params_row.Table.Columns.Contains("Organization"))
                                { userData.orgName = params_row["Organization"].ToString(); }

                                if (params_row.Table.Columns.Contains("UseVMNameComputerName")
                                    && Boolean.Parse(params_row["UseVMNameComputerName"].ToString()))
                                {
                                    userData.computerName = new CustomizationVirtualMachineName();
                                }
                                else
                                {
                                    if (params_row.Table.Columns.Contains("ComputerName"))
                                    {
                                        CustomizationFixedName name = new CustomizationFixedName();
                                        name.name = params_row["ComputerName"].ToString();
                                        userData.computerName = name;
                                    }
                                }

                                if (params_row.Table.Columns.Contains("ProductKey"))
                                { userData.productId = params_row["ProductKey"].ToString(); }

                                id.userData = userData;

                                if (params_row.Table.Columns.Contains("IncludeLicenseInformation")
                                        && Boolean.Parse(params_row["IncludeLicenseInformation"].ToString()))
                                {
                                    CustomizationLicenseFilePrintData licenseData = new CustomizationLicenseFilePrintData();
                                    if (params_row["ServerRegType"].ToString() == "perSeat")
                                    {
                                        licenseData.autoMode = CustomizationLicenseDataMode.perSeat;
                                    }
                                    else
                                    {
                                        if (params_row["ServerRegType"].ToString() == "perServer")
                                        {
                                            licenseData.autoMode = CustomizationLicenseDataMode.perServer;
                                            licenseData.autoUsers = Int32.Parse(params_row["ServerMaxConnections"].ToString());
                                            licenseData.autoUsersSpecified = true;
                                        }
                                    }
                                    id.licenseFilePrintData = licenseData;
                                }

                                if (params_row.Table.Columns.Contains("RunOnceCommand")
                                    && !String.IsNullOrEmpty(params_row["RunOnceCommand"].ToString()))
                                {
                                    CustomizationGuiRunOnce commands = new CustomizationGuiRunOnce();
                                    commands.commandList = new string[] { params_row["RunOnceCommand"].ToString() };
                                    id.guiRunOnce = commands;
                                }

                                CustomizationGuiUnattended unattended = new CustomizationGuiUnattended();
                                if (params_row.Table.Columns.Contains("AdminPassword")
                                    && !String.IsNullOrEmpty(params_row["AdminPassword"].ToString()))
                                {
                                    CustomizationPassword pass = new CustomizationPassword();
                                    pass.value = params_row["AdminPassword"].ToString();
                                    pass.plainText = true;
                                    unattended.password = pass;

                                    if (params_row.Table.Columns.Contains("AdminAutoLogin")
                                        && Boolean.Parse(params_row["AdminAutoLogin"].ToString())
                                        )
                                    {
                                        unattended.autoLogon = true;
                                        unattended.autoLogonCount = Int32.Parse(params_row["AdminAutoLoginCount"].ToString());
                                    }
                                    id.guiUnattended = unattended;
                                }
                                id.guiUnattended = unattended;

                                if (params_row.Table.Columns.Contains("DomainConfiguration"))
                                {
                                    CustomizationIdentification identification = new CustomizationIdentification();

                                    if (params_row["DomainConfiguration"].ToString() == "Workgroup")
                                    {
                                        identification.joinWorkgroup = params_row["DomainWorkgroup"].ToString();
                                    }
                                    if (params_row["DomainConfiguration"].ToString() == "Domain")
                                    {
                                        identification.joinDomain = params_row["DomainServerDomain"].ToString();
                                        identification.domainAdmin = params_row["DomainServerUsername"].ToString();
                                        CustomizationPassword pass = new CustomizationPassword();
                                        pass.value = params_row["DomainServerPassword"].ToString();
                                        pass.plainText = true;
                                        identification.domainAdminPassword = pass;
                                    }
                                    id.identification = identification;
                                }

                                it.spec.identity = id;


                                //it.spec.nicSettingMap
                                CustomizationAdapterMapping map = new CustomizationAdapterMapping();
                                map.adapter = new CustomizationIPSettings();
                                if (params_row.Table.Columns.Contains("NetworkIPConfig") && params_row["NetworkIPConfig"].ToString() == "IPSetting")
                                {
                                    CustomizationFixedIp ip = new CustomizationFixedIp();
                                    ip.ipAddress = params_row["NetworkIPConfigIP"].ToString();
                                    map.adapter.ip = ip;
                                    map.adapter.subnetMask = params_row["NetworkIPConfigSubnet"].ToString();
                                    map.adapter.gateway = new string[]{params_row["NetworkIPConfigDefaultGateway"].ToString(),
                                                                        params_row["NetworkIPConfigAlternateGateway"].ToString()};
                                }
                                else { map.adapter.ip = new CustomizationDhcpIpGenerator(); }

                                if (params_row.Table.Columns.Contains("NetworkDNSConfig") && params_row["NetworkDNSConfig"].ToString() == "DNSSetting")
                                {
                                    map.adapter.dnsServerList = new string[]{params_row["NetworkDNSConfigDefault"].ToString(),
                                                                                params_row["NetworkDNSConfigSecondary"].ToString()};
                                }
                                it.spec.nicSettingMap = new CustomizationAdapterMapping[] { map };

                                CustomizationIdentitySettings idSet = new CustomizationIdentitySettings();
                                //idSet.
                                //it.spec.identity 
                                //spec.options

                                //OS options
                                CustomizationWinOptions options = new CustomizationWinOptions();
                                if (params_row.Table.Columns.Contains("GenerateSID")
                                    && !String.IsNullOrEmpty(params_row["GenerateSID"].ToString())
                                    && info.type == "Windows")
                                {
                                    options.changeSID = Boolean.Parse(params_row["GenerateSID"].ToString());

                                }
                                else
                                {
                                    options.changeSID = false;
                                }
                                it.spec.options = options;
                                //if (params_row.Table.Columns.Contains("deleteAccounts"))
                                //{ ((CustomizationWinOptions)it.spec.options).deleteAccounts = Boolean.Parse(params_row["deleteAccounts"].ToString()); }

                                if (!exists)
                                { cb._connection._service.CreateCustomizationSpec(cb._connection._sic.customizationSpecManager, it); }
                                else
                                { cb._connection._service.OverwriteCustomizationSpec(cb._connection._sic.customizationSpecManager, it); }

                                dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                break;
                            }
                        #endregion

                        #region VM managment commands
                        case "VMClone":
                            {
                                //System.Diagnostics.Debugger.Launch();

                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm == null)
                                { throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist"); }

                                ManagedObjectReference vmFolderRef = cb.getServiceUtil().GetMoRefProp(getDataCenter(vm, ref cb), "vmFolder");

                                VirtualMachineCloneSpec cloneSpec = new VirtualMachineCloneSpec();
                                VirtualMachineRelocateSpec relocSpec = new VirtualMachineRelocateSpec();

                                cloneSpec.template = (bool)cb.getServiceUtil().GetDynamicProperty(vm, "config.template");
                                if (params_row.Table.Columns.Contains("DeployFromTemplate")
                                    && params_row["DeployFromTemplate"] != null
                                    && params_row["DeployFromTemplate"].ToString() != "{x:Null}")
                                {
                                    bool DeployFromTemplate = Boolean.Parse(params_row["DeployFromTemplate"].ToString());
                                    if (DeployFromTemplate)
                                    {
                                        cloneSpec.template = false;
                                        if (params_row.Table.Columns.Contains("VMHost")
                                            && params_row["VMHost"] != null
                                            && params_row["VMHost"].ToString() != "{x:Null}")
                                        {
                                            relocSpec.host = VMWare.Program.GetHostByName(params_row["VMHost"].ToString(), ref cb);
                                        }
                                        ManagedObjectReference resourcePool = cb.getServiceUtil().GetFirstDecendentMoRef(null, "ResourcePool");
                                        relocSpec.pool = resourcePool;
                                        //relocSpec.host //?
                                    }
                                }

                                cloneSpec.location = relocSpec;
                                cloneSpec.powerOn = false;


                                if (params_row.Table.Columns.Contains("ConfigSpec") && params_row["ConfigSpec"] != null && params_row["ConfigSpec"].ToString() != "{x:Null}")
                                {
                                    string configSpec = params_row["ConfigSpec"].ToString();

                                    if (cb._connection._service.DoesCustomizationSpecExist(cb._connection._sic.customizationSpecManager, configSpec))
                                    {
                                        cloneSpec.customization = cb._connection._service.GetCustomizationSpec(cb._connection._sic.customizationSpecManager, configSpec).spec;
                                        //cloneSpec.customization.identity.
                                        //cloneSpec.config.name
                                    }
                                    else
                                    {
                                        throw new Exception("Customization Spec " + configSpec + " doesn't exist");
                                    }
                                }

                                if (params_row.Table.Columns.Contains("ConfigSpecFile") && params_row["ConfigSpecFile"] != null && params_row["ConfigSpec"].ToString() != "{x:Null}")
                                {
                                    string configSpecFile = params_row["ConfigSpecFile"].ToString();
                                    string xml = "";
                                    try
                                    {
                                        xml = System.IO.File.ReadAllText(params_row["XML"].ToString());
                                    }
                                    catch { throw new Exception("Failed to load file " + configSpecFile); }
                                    CustomizationSpecItem it = cb._connection._service.XmlToCustomizationSpecItem(cb._connection._sic.customizationSpecManager, xml);

                                    cloneSpec.customization = it.spec;
                                }


                                /*
                                string xml = System.IO.File.ReadAllText(params_row["XML"].ToString());
                                CustomizationSpecItem it = cb._connection._service.XmlToCustomizationSpecItem(cb._connection._sic.customizationSpecManager, xml);
                                
                                bool exists = cb._connection._service.DoesCustomizationSpecExist(cb._connection._sic.customizationSpecManager,"spec1");
                                cb._connection._service.CreateCustomizationSpec(cb._connection._sic.customizationSpecManager,it);
                                CustomizationSpecItem it2 = cb._connection._service.GetCustomizationSpec(cb._connection._sic.customizationSpecManager,"spec1");

                                //    new VMware.Vim.EnvironmentBrowser(_client, new VMware.Vim.ManagedObjectReference("EnvironmentBrowser-envbrowser-7"));
                                //_this.QueryConfigOptionDescriptor();
                                
                                VirtualMachineConfigOptionDescriptor[] desc =  cb._connection._service.QueryConfigOptionDescriptor(cb._connection._service.);
                                */

                                if (params_row["CloneName"].ToString() == "") { throw new Exception("Clone name can not be empty"); }

                                var targetMachineName = params_row["CloneName"].ToString();
                                targetMachineName = TextEncode(targetMachineName);

                                ManagedObjectReference cloneTask = cb.getConnection()._service.CloneVM_Task(vm, vmFolderRef, params_row["CloneName"].ToString(), cloneSpec);
                                dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, cloneTask));
                                break;
                            }
                        case "VMCreate":
                            {
                                ManagedObjectReference rootFolderMor = cb.getConnection()._sic.rootFolder;

                                ArrayList lsDatacenters = cb.getServiceUtil().GetDecendentMoRefs(null, "Datacenter");
                                ManagedObjectReference datastoreMor = null;//cb.getServiceUtil().GetDecendentMoRef(rootFolderMor, "Datastore", params_row["DatastoreName"].ToString());
                                ManagedObjectReference datacenterMor = null;
                                foreach (Object[] o in lsDatacenters)
                                {
                                    ManagedObjectReference datacenter_mor = cb.getServiceUtil().GetDecendentMoRef(null, "Datacenter", o[1].ToString());
                                    ManagedObjectReference[] datastores_mor = (ManagedObjectReference[])cb.getServiceUtil().GetDynamicProperty(datacenter_mor, "datastore");
                                    foreach (ManagedObjectReference datastore_mor in datastores_mor)
                                    {
                                        if (cb.getServiceUtil().GetDynamicProperty(datastore_mor, "info.name").ToString() == params_row["DatastoreName"].ToString())
                                        {
                                            datastoreMor = datastore_mor;
                                            datacenterMor = datacenter_mor;
                                        }
                                    }
                                }

                                if (datacenterMor == null)
                                { throw new Exception("Datastore was not found"); }

                                ManagedObjectReference hostmor = cb.getServiceUtil().GetFirstDecendentMoRef(datacenterMor, "HostSystem");
                                ManagedObjectReference vmFolderMor = cb.getServiceUtil().GetMoRefProp(datacenterMor, "vmFolder");

                                ManagedObjectReference crmor = cb.getServiceUtil().GetMoRefProp(hostmor, "parent");
                                ManagedObjectReference resourcePool = cb.getServiceUtil().GetMoRefProp(crmor, "resourcePool");


                                VMUtils vmUtils = new VMUtils(cb);
                                int diskSize = -1;
                                if (!int.TryParse(params_row["DiskSize"].ToString(), out diskSize))
                                { throw new Exception("Invalid Disk Size parameter"); }
                                else
                                { diskSize = diskSize * 1024; }
                                //dt.Rows.Add(dt.NewRow()["Result"] = "Disc size sent:" + params_row["DiskSize"].ToString());
                                VirtualMachineConfigSpec vmConfigSpec = vmUtils.createVmConfigSpec(params_row["VMName"].ToString(),
                                                                                                   params_row["DatastoreName"].ToString(),
                                                                                                   diskSize,
                                                                                                    crmor, hostmor);

                                VirtualLsiLogicController scsiCtrl = (VirtualLsiLogicController)vmConfigSpec.deviceChange[0].device;
                                scsiCtrl.sharedBus = VirtualSCSISharing.noSharing;

                                vmConfigSpec.name = params_row["VMName"].ToString();
                                vmConfigSpec.annotation = params_row["VMAnnotation"].ToString();
                                try
                                {
                                    vmConfigSpec.memoryMB = (long)(int.Parse(params_row["MemorySize"].ToString()));
                                }
                                catch (Exception) { throw new Exception("Invalid Memory Size parameter"); }
                                vmConfigSpec.memoryMBSpecified = true;
                                try
                                {
                                    vmConfigSpec.numCPUs = int.Parse(params_row["CpuCount"].ToString());
                                }
                                catch (Exception) { throw new Exception("Invalid Cpu Count parameter"); }
                                vmConfigSpec.numCPUsSpecified = true;
                                vmConfigSpec.guestId = params_row["GuestId"].ToString();

                                ManagedObjectReference taskmor = cb.getConnection()._service.CreateVM_Task(
                                        vmFolderMor, vmConfigSpec, resourcePool, hostmor
                                );
                                dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));

                                break;
                            }
                        case "VMDelete":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection()._service.Destroy_Task(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        #endregion

                        #region Power commands
                        case "VMPowerState":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    VirtualMachineRuntimeInfo vmRuntimeInfo = (VirtualMachineRuntimeInfo)cb.getServiceUtil().GetDynamicProperty(vm, "runtime");
                                    VirtualMachinePowerState state = vmRuntimeInfo.powerState;
                                    dt.Rows.Add(dt.NewRow()["Result"] = ConvertPowerStateString(state));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMPowerOFF":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection().Service.PowerOffVM_Task(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMPowerON":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection().Service.PowerOnVM_Task(vm, null);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMSuspend":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection().Service.SuspendVM_Task(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMReset":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection().Service.ResetVM_Task(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMReboot":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    cb.getConnection().Service.RebootGuest(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMShutDown":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    cb.getConnection().Service.ShutdownGuest(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMStandby":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    cb.getConnection().Service.StandbyGuest(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        #endregion

                        #region Snapshot commands
                        case "VMCreateSnapshot":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection()._service.CreateSnapshot_Task(
                                        vm,
                                        params_row["SnapshotName"].ToString(),
                                        params_row["SnapshotDescription"].ToString(),
                                        Boolean.Parse(params_row["MemoryStateDump"].ToString()),
                                        Boolean.Parse(params_row["PauseDiskWrites"].ToString()));
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMListSnapshot":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    VirtualMachineSnapshotTree[] snt = GetSnapshotTree(cb, vm);
                                    traverseSnapshotInTree(snt, "", ref dt);
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMCurrentSnapshot":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    object o = cb.getServiceUtil().GetDynamicProperty(vm, "snapshot");
                                    if (o != null)
                                    {
                                        VirtualMachineSnapshotInfo inf = (VirtualMachineSnapshotInfo)o;
                                        VirtualMachineSnapshotTree[] snt = GetSnapshotTree(cb, vm);
                                        VirtualMachineSnapshotTree t = traverseSnapshotInTree(snt, inf.currentSnapshot);

                                        dt.Rows.Add(dt.NewRow()["Result"] = t.name);

                                    }
                                    else
                                    {
                                        dt.Rows.Add(dt.NewRow()["Result"] = "No snapshots");
                                    }
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMDeleteSnapshot":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    VirtualMachineSnapshotTree[] snt = GetSnapshotTree(cb, vm);
                                    VirtualMachineSnapshotTree sn_mor = traverseSnapshotInTree(snt, params_row["SnapshotName"].ToString(), ref dt);
                                    if (sn_mor != null)
                                    {
                                        ManagedObjectReference taskmor = cb.getConnection()._service.RemoveSnapshot_Task(sn_mor.snapshot, Boolean.Parse(params_row["DeleteChildren"].ToString()), true);
                                        dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                    }
                                    else
                                    {
                                        //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["SnapshotName"].ToString() + " doesn't exist");
                                        throw new Exception(params_row["SnapshotName"].ToString() + " doesn't exist");
                                    }
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMDeleteAllSnapshots":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection()._service.RemoveAllSnapshots_Task(vm, true);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMRenameSnapshot":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    VirtualMachineSnapshotTree[] snt = GetSnapshotTree(cb, vm);
                                    VirtualMachineSnapshotTree sn_mor = traverseSnapshotInTree(snt, params_row["SnapshotName"].ToString(), ref dt);
                                    if (sn_mor != null)
                                    {
                                        cb.getConnection()._service.RenameSnapshot(sn_mor.snapshot, params_row["NewSnapshotName"].ToString(), sn_mor.description);
                                        dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                    }
                                    else
                                    {
                                        //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["SnapshotName"].ToString() + " doesn't exist");
                                        throw new Exception("Snapshot " + params_row["SnapshotName"].ToString() + " doesn't exist");
                                    }
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMSnapshotInfo":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    VirtualMachineSnapshotTree[] snt = GetSnapshotTree(cb, vm);
                                    VirtualMachineSnapshotTree sn_mor = traverseSnapshotInTree(snt, params_row["SnapshotName"].ToString(), ref dt);
                                    if (sn_mor != null)
                                    {

                                        DataTable dtResult = new DataTable("resultSet");
                                        dtResult.Columns.Add("Name", typeof(String));
                                        dtResult.Columns.Add("VM name", typeof(String));
                                        dtResult.Columns.Add("Description", typeof(String));
                                        dtResult.Columns.Add("State", typeof(String));
                                        dtResult.Columns.Add("Quiesced", typeof(String));
                                        dtResult.Columns.Add("Creation time", typeof(String));

                                        DataRow r = dtResult.NewRow();
                                        r["Name"] = sn_mor.name;
                                        r["VM name"] = cb.getServiceUtil().GetDynamicProperty(sn_mor.vm, "config.name").ToString();
                                        r["Description"] = sn_mor.description;
                                        r["State"] = ConvertPowerStateString(sn_mor.state);
                                        r["Quiesced"] = sn_mor.quiesced.ToString();
                                        r["Creation time"] = sn_mor.createTime.ToString();

                                        dtResult.Rows.Add(r);
                                        dt = dtResult;
                                    }
                                    else
                                    {
                                        //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["SnapshotName"].ToString() + " doesn't exist");
                                        throw new Exception("Snapshot " + params_row["SnapshotName"].ToString() + " doesn't exist");
                                    }
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMRevertToCurrentSnapshot":
                            {
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    //ManagedObjectReference host = GetHostForVM(vm, ref cb);
                                    //object o = cb.getServiceUtil().GetDynamicProperty(vm, "snapshot");
                                    //if (o != null)
                                    //{
                                    // VirtualMachineSnapshotInfo inf = (VirtualMachineSnapshotInfo)o;
                                    ManagedObjectReference taskmor = cb._connection._service.RevertToCurrentSnapshot_Task(vm, null, true);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));

                                    // }
                                    //else
                                    // {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - No snapshots");
                                    // throw new Exception("No snapshots");
                                    // }
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMRevertToSnapshot":
                            {
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    VirtualMachineSnapshotTree[] snt = GetSnapshotTree(cb, vm);
                                    VirtualMachineSnapshotTree sn_mor = traverseSnapshotInTree(snt, params_row["SnapshotName"].ToString(), ref dt);
                                    if (sn_mor != null)
                                    {
                                        ManagedObjectReference host = GetHostForVM(vm, ref cb);
                                        ManagedObjectReference taskmor = cb._connection._service.RevertToSnapshot_Task(sn_mor.snapshot, host, true);
                                        dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));

                                    }
                                    else
                                    {
                                        //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["SnapshotName"].ToString() + " doesn't exist");
                                        throw new Exception("Snapshot" + params_row["SnapshotName"].ToString() + " doesn't exist");
                                    }
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        #endregion

                        #region Template commands
                        case "VMMarkTemplate":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    cb.getConnection()._service.MarkAsTemplate(vm);
                                    dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMUnMarkTemplate":
                            {
                                //ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", params_row["VMName"].ToString());
                                ManagedObjectReference vm = GetVmMorFromNameOrUuid(ref cb, params_row["VMName"].ToString());
                                if (vm != null)
                                {
                                    ManagedObjectReference resourcePool = null;
                                    ManagedObjectReference hostName = null;
                                    //ManagedObjectReference resourcePool = GetNamedOrDefaultResourcePool(params_row["ResourcePool"].ToString(),ref cb);
                                    if (params_row["ResourcePool"] != null && params_row["ResourcePool"].ToString() != "")
                                    {
                                        resourcePool = cb.getServiceUtil().GetDecendentMoRef(null, "ResourcePool", params_row["ResourcePool"].ToString());
                                    }
                                    else
                                    {
                                        if (params_row["HostSystemName"] != null && params_row["HostSystemName"].ToString() != "")
                                        { hostName = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["HostSystemName"].ToString()); }

                                        if (hostName != null)
                                        {
                                            ManagedObjectReference crmor = cb.getServiceUtil().GetMoRefProp(hostName, "parent");
                                            resourcePool = cb.getServiceUtil().GetMoRefProp(crmor, "resourcePool");
                                        }
                                        else
                                        { throw new Exception("Hostname is incorrect"); }
                                    }

                                    //dt.Rows.Add(dt.NewRow()["Result"] = "res_pool" + resourcePool.Value);
                                    cb.getConnection()._service.MarkAsVirtualMachine(vm, resourcePool, hostName);
                                    dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                                    throw new Exception("VM " + params_row["VMName"].ToString() + " doesn't exist");
                                }
                                break;
                            }
                        case "VMListTemplates":
                            {
                                System.Collections.ArrayList vms = cb.getServiceUtil().GetDecendentMoRefs(null, "VirtualMachine");

                                foreach (Object[] vm_ref in vms)
                                {
                                    ManagedObjectReference vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", vm_ref[1].ToString());
                                    object template_prop = cb.getServiceUtil().GetDynamicProperty(vm, "config.template");
                                    if (template_prop != null && (bool)template_prop)
                                    {
                                        dt.Rows.Add(dt.NewRow()["Result"] = vm_ref[1].ToString());
                                    }
                                }
                                //Sort
                                if (dt.Rows.Count > 0)
                                {
                                    DataView dv = dt.DefaultView;
                                    dv.Sort = "Result";
                                    dt = dv.ToTable();
                                }
                                break;
                            }
                        #endregion

                        #region Host commands
                        case "VMHostReboot":
                            {
                                ManagedObjectReference comp_mor = cb.getServiceUtil().GetDecendentMoRef(null, "ComputeResource", params_row["TargetHostName"].ToString());

                                ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["TargetHostName"].ToString());
                                if (host_mor != null)
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Host:" + host_mor);
                                    ManagedObjectReference taskmor = cb.getConnection()._service.RebootHost_Task(host_mor, bool.Parse(params_row["ForceReboot"].ToString()));
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Target host name not valid");
                                    throw new Exception("Target host name not valid");
                                }
                                break;
                            }
                        case "VMHostEnterMaintenanceMode":
                            {
                                ManagedObjectReference comp_mor = cb.getServiceUtil().GetDecendentMoRef(null, "ComputeResource", params_row["TargetHostName"].ToString());

                                ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["TargetHostName"].ToString());
                                if (host_mor != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection()._service.EnterMaintenanceMode_Task(host_mor, int.Parse(params_row["Timeout"].ToString()), true, null);
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Target host name not valid");
                                    throw new Exception("Target host name not valid");
                                }
                                break;
                            }
                        case "VMHostExitMaintenanceMode":
                            {
                                ManagedObjectReference comp_mor = cb.getServiceUtil().GetDecendentMoRef(null, "ComputeResource", params_row["TargetHostName"].ToString());

                                ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["TargetHostName"].ToString());

                                if (host_mor != null)
                                {
                                    ManagedObjectReference taskmor = cb.getConnection()._service.ExitMaintenanceMode_Task(host_mor, int.Parse(params_row["Timeout"].ToString()));
                                    dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Target host name not valid");
                                    throw new Exception("Target host name not valid");
                                }
                                break;
                            }
                        #endregion

                        #region File commands
                        //case "VMCopyFile":
                        //    {

                        //        ManagedObjectReference comp_mor = cb.getServiceUtil().GetDecendentMoRef(null, "ComputeResource", params_row["TargetHostName"].ToString());

                        //        ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["TargetHostName"].ToString());

                        //        ManagedObjectReference taskmor = cb.getConnection()._service.ExitMaintenanceMode_Task(host_mor, int.Parse(params_row["Timeout"].ToString()));
                        //        dt.Rows.Add(dt.NewRow()["Result"] = GetTaskResult(cb, taskmor));
                        //        break;
                        //    }
                        case "VMDeleteFile":
                            {
                                //cb.getConnection()._service.DeleteFile(
                                break;
                            }
                        case "VMMoveFile":
                            {
                                //cb.getConnection()._service.MoveIntoFolder_Task(
                                break;
                            }
                        case "VMFolderCreate":
                            {

                                ManagedObjectReference folderMoRef = cb.getServiceUtil().GetDecendentMoRef(null, "Folder", params_row["ParentFolder"].ToString());
                                if (folderMoRef != null)
                                {
                                    cb.getConnection()._service.CreateFolder(folderMoRef, params_row["NewFolder"].ToString()); // returns the actual folder mor
                                    dt.Rows.Add(dt.NewRow()["Result"] = "Success");
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Folder " + params_row["ParentFolder"].ToString() + " not found");
                                    throw new Exception("Folder " + params_row["ParentFolder"].ToString() + " not found");
                                }
                                break;
                            }
                        #endregion

                        #region Perf Commands
                        case "VMHostGetCpuUsage":
                            {
                                string groupName = "CPU";
                                string name = "usage";
                                string rollup = "average";
                                string targetHostName = params_row["TargetHostName"].ToString();

                                dt = GetCounterMetrics(ref cb, groupName, name, rollup, targetHostName);
                                break;
                            }
                        case "VMHostGetMemoryUsage":
                            {
                                string groupName = "Memory";
                                string name = "usage";
                                string rollup = "average";
                                string targetHostName = params_row["TargetHostName"].ToString();

                                dt = GetCounterMetrics(ref cb, groupName, name, rollup, targetHostName);
                                break;
                            }
                        case "VMHostGetStorageUsage":
                            {
                                break;
                            }
                        case "VMHostGetState":
                            {
                                ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["TargetHostName"].ToString());
                                if (host_mor != null)
                                {
                                    HostRuntimeInfo hsrt = (HostRuntimeInfo)cb.getServiceUtil().GetDynamicProperty(host_mor, "runtime");
                                    string state = hsrt.connectionState.ToString();
                                    if (hsrt.inMaintenanceMode)
                                    { state += " - In maintenance mode"; }
                                    dt.Rows.Add(dt.NewRow()["Result"] = state);
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Target host name not valid");
                                    throw new Exception("Target host name not valid");
                                }
                                break;
                            }
                        case "VMPerfCounterList":
                            {
                                //System.Diagnostics.Debugger.Launch();
                                ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", params_row["TargetHostName"].ToString());
                                if (host_mor != null)
                                {
                                    //int[] ids = new int[1000];
                                    //for (int i = 0; i < 1000; ++i) {
                                    //    ids[i] = i;
                                    //}
                                    ManagedObjectReference perf = cb._connection._sic.perfManager;
                                    PerfCounterInfo[] counters = (PerfCounterInfo[])cb.getServiceUtil().GetDynamicProperty(perf, "perfCounter");
                                    //PerfCounterInfo[] counters = cb._connection._service.QueryPerfCounter(perf, ids);

                                    DataTable dtResult = new DataTable("resultSet");
                                    dtResult.Columns.Add("Counter key", typeof(String));
                                    //dtResult.Columns.Add("Group Info Key", typeof(String));
                                    dtResult.Columns.Add("Group Info Label", typeof(String));
                                    //dtResult.Columns.Add("Group Info Summary", typeof(String));

                                    dtResult.Columns.Add("Name", typeof(String));
                                    //dtResult.Columns.Add("Name Info Label", typeof(String));
                                    dtResult.Columns.Add("Description", typeof(String));
                                    dtResult.Columns.Add("Rollup", typeof(String));
                                    dtResult.Columns.Add("Stats type", typeof(String));
                                    //dtResult.Columns.Add("Unit Info Key", typeof(String));
                                    //dtResult.Columns.Add("Unit Info Label", typeof(String));
                                    dtResult.Columns.Add("Unit", typeof(String));

                                    for (int i = 0; i < counters.Length; ++i)
                                    {
                                        DataRow r = dtResult.NewRow();

                                        r["Counter key"] = counters[i].key.ToString();
                                        //r["Group Info Key"] = counters[i].groupInfo.key;
                                        r["Group Info Label"] = counters[i].groupInfo.label;
                                        //["Group Info Summary"] = counters[i].groupInfo.summary;

                                        r["Name"] = counters[i].nameInfo.key;
                                        //r["Name Info Label"] = counters[i].nameInfo.label;
                                        r["Description"] = counters[i].nameInfo.summary;
                                        r["Rollup"] = counters[i].rollupType.ToString();
                                        r["Stats type"] = counters[i].statsType.ToString();
                                        //r["Unit Info Key"] = counters[i].unitInfo.key;
                                        //r["Unit Info Label"] = counters[i].unitInfo.label;
                                        r["Unit"] = counters[i].unitInfo.summary;


                                        dtResult.Rows.Add(r);
                                    }

                                    dt = dtResult;
                                }
                                else
                                {
                                    //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Target host name not valid");
                                    throw new Exception("Target host name not valid");
                                }
                                break;
                            }
                        case "VMPerfGetCounterReading":
                            {
                                //System.Diagnostics.Debugger.Launch();
                                string groupName = params_row["GroupName"].ToString();
                                string name = params_row["Name"].ToString();
                                string rollup = params_row["Rollup"].ToString();
                                string targetHostName = params_row["TargetHostName"].ToString();

                                dt = GetCounterMetrics(ref cb, groupName, name, rollup, targetHostName);
                                break;
                            }
                        #endregion
                        //default: dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Unknown command"); break;
                        default: throw new Exception("Unknown command"); //break;
                    }
                    //}
                    //else
                    //{
                    //    dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + params_row["VMName"].ToString() + " doesn't exist");
                    //}

                    cb.disConnect();

                    if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\n Connection to ESXi/VSphere server closed"); }
                }
            }
            catch (Exception e)
            {
                //dt = new DataTable("resultSet");
                //dt.Columns.Add("Result");
                //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - " + e.Message);
                string message = e.Message;

                message = message.Replace("\n", "");
                //truncate stack traces if pressent
                int i = message.IndexOf("while parsing call information for method");
                if (i > -1)
                {
                    message = message.Substring(0, i);
                }

                dt = new DataTable("resultSet");
                dt.Columns.Add("error");
                dt.Columns.Add("errorstack");
                dt.Columns.Add("errorinner");

                DataRow r = dt.NewRow();
                r["error"] = message;
                r["errorstack"] = e.StackTrace;
                if (e.InnerException != null)
                { r["errorinner"] = e.InnerException.ToString(); }
                else
                { r["errorinner"] = ""; }
                dt.Rows.Add(r);
            }

            if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\n Writing results"); }

            dt.WriteXml(sw, XmlWriteMode.WriteSchema, false);
            if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\n Result is: " + sw.ToString()); }

            Console.WriteLine(sw.ToString());
            if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\n Result is written to console"); }
            if (File.Exists(@"C:\Temp\Vmlog.log")) { System.IO.File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\n Application exit"); }
        }

        private static string GetFolder(ref AppUtil.AppUtil cb, ManagedObjectReference vm)
        {
            string sbPath = "";
            ManagedObjectReference mor = (ManagedObjectReference)cb.getServiceUtil().GetDynamicProperty(vm, "parent");
            if (mor != null && mor.type == "Folder")
            {
                //sbPath = mor.Value;
                sbPath = (string)cb.getServiceUtil().GetDynamicProperty(mor, "name");
            }
            return sbPath;

        }

        private static string GetFolderPath(ref AppUtil.AppUtil cb, ManagedObjectReference vm)
        {
            string sbPath = "";
            ManagedObjectReference mor = vm;
            while (mor != null)
            {
                //Datastore
                if (mor != null && mor.type == "Datacenter")
                {
                    sbPath = "[" + mor.Value + "]\\" + sbPath;
                }
                if (mor != null && mor.type == "Folder")
                {
                    sbPath = mor.Value + "\\" + sbPath;
                }
                mor = (ManagedObjectReference)cb.getServiceUtil().GetDynamicProperty(mor, "parent");
            }
            return sbPath;

        }

        private static VirtualDeviceConfigSpec getDiskDeviceConfigSpec(ref AppUtil.AppUtil cb, ManagedObjectReference _virtualMachine)
        {
            String ops = cb.get_option("operation");
            VirtualDeviceConfigSpec diskSpec = new VirtualDeviceConfigSpec();
            VirtualMachineConfigInfo vmConfigInfo
               = (VirtualMachineConfigInfo)cb.getServiceUtil().GetDynamicProperty(
                   _virtualMachine, "config");

            if (ops.Equals("add"))
            {
                VirtualDisk disk = new VirtualDisk();
                VirtualDiskFlatVer2BackingInfo diskfileBacking
                   = new VirtualDiskFlatVer2BackingInfo();
                String dsName = "TODO";
                //String dsName = getDataStoreName(int.Parse(cb.get_option("disksize")));

                int ckey = -1;
                int unitNumber = 0;

                VirtualDevice[] test = vmConfigInfo.hardware.device;
                for (int k = 0; k < test.Length; k++)
                {
                    if (test[k].deviceInfo.label.Equals(
                       "SCSI Controller 0"))
                    {
                        ckey = test[k].key;
                    }
                }




                unitNumber = test.Length + 1;
                String fileName = "[" + dsName + "] " + cb.get_option("vmname")
                                + "/" + cb.get_option("value") + ".vmdk";

                diskfileBacking.fileName = fileName;
                diskfileBacking.diskMode = cb.get_option("diskmode");

                disk.controllerKey = ckey;
                disk.unitNumber = unitNumber;
                disk.controllerKeySpecified = true;
                disk.unitNumberSpecified = true;
                disk.backing = diskfileBacking;
                int size = 1024 * (int.Parse(cb.get_option("disksize")));
                disk.capacityInKB = size;
                disk.key = 0;

                diskSpec.operation = VirtualDeviceConfigSpecOperation.add;
                diskSpec.fileOperation = VirtualDeviceConfigSpecFileOperation.create;
                diskSpec.fileOperationSpecified = true;
                diskSpec.operationSpecified = true;
                diskSpec.device = disk;
            }
            else if (ops.Equals("remove"))
            {
                VirtualDisk disk = null;
                VirtualDiskFlatVer2BackingInfo diskfileBacking
                   = new VirtualDiskFlatVer2BackingInfo();

                VirtualDevice[] test = vmConfigInfo.hardware.device;
                for (int k = 0; k < test.Length; k++)
                {
                    if (test[k].deviceInfo.label.Equals(
                            cb.get_option("value")))
                    {
                        disk = (VirtualDisk)test[k];
                    }
                }
                if (disk != null)
                {
                    diskSpec.operation = VirtualDeviceConfigSpecOperation.remove;
                    diskSpec.operationSpecified = true;
                    diskSpec.fileOperation = VirtualDeviceConfigSpecFileOperation.destroy;
                    diskSpec.fileOperationSpecified = true;
                    diskSpec.device = disk;
                }
                else
                {
                    Console.WriteLine("No device found " + cb.get_option("value"));
                    return null;
                }
            }
            return diskSpec;
        }

        private static DataTable GetCounterMetrics(ref AppUtil.AppUtil cb, string groupName, string name, string rollup, string targetHostName)
        {
            DataTable dt = new DataTable("resultSet");
            dt.Columns.Add("Result", typeof(String));

            ManagedObjectReference host_mor = cb.getServiceUtil().GetDecendentMoRef(null, "HostSystem", targetHostName);
            if (host_mor != null)
            {
                ManagedObjectReference perf = cb._connection._sic.perfManager;

                int IntervalId = cb._connection._service.QueryPerfProviderSummary(perf, host_mor).refreshRate;
                PerfCounterInfo[] counters = (PerfCounterInfo[])cb.getServiceUtil().GetDynamicProperty(perf, "perfCounter");

                if (!String.IsNullOrEmpty(groupName))
                { counters = counters.Where(c => c.groupInfo.label == groupName).ToArray<PerfCounterInfo>(); }

                if (!String.IsNullOrEmpty(name))
                { counters = counters.Where(c => c.nameInfo.key == name).ToArray<PerfCounterInfo>(); }

                if (!String.IsNullOrEmpty(rollup))
                { counters = counters.Where(c => c.rollupType.ToString() == rollup).ToArray<PerfCounterInfo>(); }

                //ILookup<int,PerfCounterInfo> lCounters = counters.ToLookup(m => m.key);
                IDictionary dCounters = counters.ToDictionary(m => m.key);
                QueryAvailablePerfMetricRequest request = new QueryAvailablePerfMetricRequest(perf, host_mor, DateTime.Now, DateTime.Now, IntervalId);
                // PerfMetricId[] metricIds = cb._connection._service.QueryAvailablePerfMetric(perf, host_mor, DateTime.Now, false, DateTime.Now, false, IntervalId, true);
                PerfMetricId[] metricIds = cb._connection._service.QueryAvailablePerfMetric(request).returnval;
                metricIds = metricIds.Where(m => dCounters.Contains(m.counterId)).ToArray<PerfMetricId>();

                PerfQuerySpec qSpec = new PerfQuerySpec();
                qSpec.entity = host_mor;
                qSpec.maxSample = 1;
                qSpec.maxSampleSpecified = true;
                qSpec.metricId = metricIds;
                qSpec.intervalId = IntervalId;
                qSpec.intervalIdSpecified = true;
                PerfQuerySpec[] qSpecs = new PerfQuerySpec[] { qSpec };
                QueryPerfRequest queryPerfRequest = new QueryPerfRequest(perf, qSpecs);
                //PerfEntityMetricBase[] emb = cb._connection._service.QueryPerf(perf, qSpecs);
                PerfEntityMetricBase[] emb = cb._connection._service.QueryPerf(queryPerfRequest).returnval;

                DataTable dtResult = new DataTable("resultSet");
                dtResult.Columns.Add("Instance", typeof(String));
                dtResult.Columns.Add("Value", typeof(String));

                for (int i = 0; i < emb.Length; ++i)
                {
                    PerfMetricSeries[] vals = ((PerfEntityMetric)emb[i]).value;
                    PerfSampleInfo[] infos = ((PerfEntityMetric)emb[i]).sampleInfo;
                    //Console.WriteLine("Sample time range: " +
                    //                  infos[0].timestamp.TimeOfDay.ToString() + " - " +
                    //                  infos[infos.Length - 1].timestamp.TimeOfDay.ToString());
                    for (int vi = 0; vi < vals.Length; vi++)
                    {
                        PerfCounterInfo pci = (PerfCounterInfo)dCounters[vals[vi].id.counterId];
                        //if (pci != null)
                        //    Console.WriteLine(pci.nameInfo.summary);
                        if (vals[vi].GetType().Name.Equals("PerfMetricIntSeries"))
                        {
                            PerfMetricIntSeries val = (PerfMetricIntSeries)vals[vi];
                            DataRow r = dtResult.NewRow();

                            r["Instance"] = val.id.instance;

                            long[] longs = val.value;
                            string sResult = "";
                            for (int k = 0; k < longs.Length; ++k)
                            {
                                if (pci.unitInfo.summary == "Percentages")
                                { sResult += ((float)longs[k] / (float)100) + " "; }
                                else { sResult += longs[k] + " "; }
                            }
                            r["Value"] = sResult;

                            dtResult.Rows.Add(r);
                            //long[] longs = val.value;
                            //for (int k = 0; k < longs.Length; ++k)
                            //{
                            //    Console.WriteLine(longs[k] + " ");
                            //}
                            //Console.WriteLine();
                        }
                        else
                        {
                            dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Unexpected metric format");
                        }
                    }
                }
                dt = dtResult;
            }
            else
            {
                //dt.Rows.Add(dt.NewRow()["Result"] = "Failure - Target host name not valid");
                throw new Exception("Target host name not valid");
            }
            return dt;
        }


        private static DatastoreSummary getDataStoreSummary(ManagedObjectReference vm, ref AppUtil.AppUtil cb)
        {
            DatastoreSummary dsSum = null;
            //VirtualMachineRuntimeInfo vmRuntimeInfo
            //   = (VirtualMachineRuntimeInfo)cb.getServiceUtil().GetDynamicProperty(
            //       _virtualMachine, "runtime");
            ManagedObjectReference envBrowser = (ManagedObjectReference)cb.getServiceUtil().GetDynamicProperty(vm, "environmentBrowser");
            //ManagedObjectReference hmor = vmRuntimeInfo.host;
            ManagedObjectReference hmor = GetHostForVM(vm, ref cb);

            if (hmor == null)
            { throw new Exception("No Datastore found for vm:" + vm.Value); }

            ConfigTarget configTarget = cb.getConnection()._service.QueryConfigTarget(envBrowser, null);
            if (configTarget.datastore != null)
            {
                for (int i = 0; i < configTarget.datastore.Length; i++)
                {
                    VirtualMachineDatastoreInfo vdsInfo = configTarget.datastore[i];
                    DatastoreSummary dsSummary = vdsInfo.datastore;
                    if (dsSummary.accessible)
                    {
                        dsSum = dsSummary;
                        break;
                    }
                }
            }
            return dsSum;
        }

        private static ManagedObjectReference GetHostForVM(ManagedObjectReference vm, ref AppUtil.AppUtil cb)
        {
            //TODO

            //    VirtualMachineRuntimeInfo vmRuntimeInfo = (VirtualMachineRuntimeInfo)cb.getServiceUtil().GetDynamicProperty(vm,"runtime");       
            //    ManagedObjectReference hmor = vmRuntimeInfo.host;

            string uuid = cb.getServiceUtil().GetDynamicProperty(vm, "config.uuid").ToString();

            ArrayList hosts = cb.getServiceUtil().GetDecendentMoRefs(null, "HostSystem");
            ManagedObjectReference result = null;
            foreach (object o_host in hosts)
            {
                ManagedObjectReference host = (ManagedObjectReference)((object[])o_host)[0];
                Object o_vm = cb.getServiceUtil().GetDecendentMoRefs(host, "VirtualMachine", new string[][] { new string[] { "config.uuid", uuid } });

                if (o_vm != null)
                {
                    result = host;
                }
            }

            return result;
        }

        private static ManagedObjectReference GetHostByName(string Name, ref AppUtil.AppUtil cb)
        {
            //TODO

            //    VirtualMachineRuntimeInfo vmRuntimeInfo = (VirtualMachineRuntimeInfo)cb.getServiceUtil().GetDynamicProperty(vm,"runtime");       
            //    ManagedObjectReference hmor = vmRuntimeInfo.host;

            ArrayList hosts = cb.getServiceUtil().GetDecendentMoRefs(null, "HostSystem");
            ManagedObjectReference result = null;
            foreach (object o_host in hosts)
            {
                ManagedObjectReference host = (ManagedObjectReference)((object[])o_host)[0];
                string host_name = cb.getServiceUtil().GetDynamicProperty(host, "name").ToString();
                if (host_name == Name)
                {
                    result = host;
                    break;
                }
            }

            return result;
        }

        private static ManagedObjectReference GetVmMorFromNameOrUuid(ref AppUtil.AppUtil cb, string VMId)
        {
            ManagedObjectReference vm = null;
            if (VMId.Length == 36 && VMId[8] == '-' && VMId[13] == '-' && VMId[18] == '-' && VMId[23] == '-')
            {
                ArrayList list = cb.getServiceUtil().GetDecendentMoRefs(null, "VirtualMachine", new string[][] { new string[] { "config.uuid", VMId } });
                if (list.Count == 1)
                {
                    vm = (ManagedObjectReference)list[0];
                }
            }
            else
            {
                vm = cb.getServiceUtil().GetDecendentMoRef(null, "VirtualMachine", VMId);
            }
            return vm;
        }

        private static string ConvertPowerStateString(VirtualMachinePowerState state)
        {
            return state.ToString().Replace("powered", "").Replace("suspended", "Suspended");
        }

        private static ManagedObjectReference getDataCenter(ManagedObjectReference vmmor, ref AppUtil.AppUtil cb)
        //private static string getDataCenter(ManagedObjectReference vmmor, ref AppUtil.AppUtil cb)
        {
            ManagedObjectReference morParent = cb.getServiceUtil().GetMoRefProp(vmmor, "parent");
            morParent = cb.getServiceUtil().GetMoRefProp(morParent, "parent");
            if (!morParent.type.Equals("Datacenter"))
            {
                morParent = cb.getServiceUtil().GetMoRefProp(morParent, "parent");
            }
            //return (ManagedObjectReference)cb.getServiceUtil().GetDynamicProperty(morParent, "name");
            Object objdcName = cb.getServiceUtil().GetDynamicProperty(morParent, "name");
            //String dcName = objdcName.ToString();
            //return dcName;
            return cb.getServiceUtil().GetDecendentMoRef(null, "Datacenter", objdcName.ToString());
        }

        private static VirtualMachineSnapshotTree[] GetSnapshotTree(AppUtil.AppUtil cb, ManagedObjectReference vm)
        {
            VirtualMachineSnapshotTree[] sntResult = null;
            ObjectContent[] snaps = cb.getServiceUtil().GetObjectProperties(null, vm, new String[] { "snapshot" });
            VirtualMachineSnapshotInfo snapInfo = null;
            if (snaps != null && snaps.Length > 0)
            {
                ObjectContent snapobj = snaps[0];
                DynamicProperty[] snapary = snapobj.propSet;
                if (snapary != null && snapary.Length > 0)
                {
                    snapInfo = ((VirtualMachineSnapshotInfo)(snapary[0]).val);
                    //VirtualMachineSnapshotTree[] snapTree = snapInfo.rootSnapshotList;
                    sntResult = snapInfo.rootSnapshotList;
                    //traverseSnapshotInTree(snapTree, null, ref dt);
                }
            }
            return sntResult;
        }

        //private static ManagedObjectReference traverseSnapshotInTree(VirtualMachineSnapshotTree[] snapTree, String findName,ref DataTable dt) 
        private static VirtualMachineSnapshotTree traverseSnapshotInTree(VirtualMachineSnapshotTree[] snapTree, String findName, ref DataTable dt)
        {
            VirtualMachineSnapshotTree snapmor = null;
            if (snapTree == null)
            {
                return snapmor;
            }
            for (int i = 0; i < snapTree.Length && snapmor == null; i++)
            {
                VirtualMachineSnapshotTree node = snapTree[i];
                if (findName == null || findName == "")
                {
                    if (dt != null)
                    {
                        dt.Rows.Add(dt.NewRow()["Result"] = node.name);
                    }
                }

                if (findName != null && node.name.Equals(findName))
                {
                    //snapmor = node.snapshot;
                    snapmor = node;
                }
                else
                {
                    VirtualMachineSnapshotTree[] childTree = node.childSnapshotList;
                    snapmor = traverseSnapshotInTree(childTree, findName, ref dt);
                }
            }

            return snapmor;
        }

        private static VirtualMachineSnapshotTree traverseSnapshotInTree(VirtualMachineSnapshotTree[] snapTree, ManagedObjectReference snapshot)
        {
            VirtualMachineSnapshotTree snapmor = null;
            if (snapTree == null)
            {
                return snapmor;
            }
            for (int i = 0; i < snapTree.Length && snapmor == null; i++)
            {
                VirtualMachineSnapshotTree node = snapTree[i];
                if (snapshot != null && node.snapshot.Value == snapshot.Value)
                {
                    //snapmor = node.snapshot;
                    snapmor = node;
                }
                else
                {
                    VirtualMachineSnapshotTree[] childTree = node.childSnapshotList;
                    snapmor = traverseSnapshotInTree(childTree, snapshot);
                }
            }

            return snapmor;
        }

        private static string GetTaskResult(AppUtil.AppUtil connection, ManagedObjectReference taskmor)
        {
            string returnResult = string.Empty;

            try
            {

                for (int iCounter = 0; iCounter < int.MaxValue; iCounter++)
                {
                    returnResult = QueryTask(connection, taskmor);

                    if (string.IsNullOrEmpty(returnResult))
                    {
                        // Wait for next turn
                        System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception exp)
            {
                if (File.Exists(@"C:\Temp\Vmlog.log")) { File.AppendAllText(@"C:\Temp\Vmlog.log", "\r\n Exception occurred while trying to query task results"); }
                throw exp;
            }

            return returnResult;
        }

        private static string QueryTask(AppUtil.AppUtil connection, ManagedObjectReference taskmor)
        {
            string returnResult = string.Empty;

            var result = WaitForTaskResult(connection, taskmor);
            if (result == null)
            {
                // Result returned but without data
                return "Error in reading task result";
            }

            TaskInfoState enumResult;

            if (Enum.TryParse<TaskInfoState>(result.ToString(), out enumResult) == true)
            {
                switch (enumResult)
                {
                    case TaskInfoState.error:
                        HandleError(connection, taskmor);
                        break;

                    case TaskInfoState.success:
                        returnResult = "Success";
                        break;

                    case TaskInfoState.queued:
                    case TaskInfoState.running:
                    default:
                        // Do nothing
                        break;
                }

                return returnResult;
            }
            else
            {
                throw new ApplicationException("Could not convert to enum");
            }

        }

        private static void HandleError(AppUtil.AppUtil connection, ManagedObjectReference taskmor)
        {
            TaskInfo tinfo = (TaskInfo)connection.getServiceUtil().GetDynamicProperty(taskmor, "info");
            LocalizedMethodFault fault = tinfo.error;
            if (fault != null)
            {
                //Console.WriteLine("Fault " + fault.fault.ToString());
                //Console.WriteLine("Message " + fault.localizedMessage);
                //return "Failure - " + fault.localizedMessage;
                throw new Exception(fault.localizedMessage);
            }
            else
            {
                //return "Failure - " + result[1].ToString();
                throw new Exception("Error occurred while trying to received information about this task");
            }
        }

        private static object WaitForTaskResult(AppUtil.AppUtil cb, ManagedObjectReference taskmor)
        {
            object[] result = cb.getServiceUtil().WaitForValues(
                                                                    taskmor,
                                                                    new string[] { "info.state", "info.error" },
                                                                    new string[] { "state" }, // info has a property - state for state of the task
                                                                    new object[][] { new object[] { TaskInfoState.success, TaskInfoState.error, TaskInfoState.running } }
                                                                    );

            if ((result != null) && (result.Length >= 1))
            {
                return result[0];
            }

            return null;
        }

        private readonly static string reservedCharacters = "!*'();:@&=+$,/?%#[]";
        private static string TextEncode(string value)
        {
            if (String.IsNullOrEmpty(value))
                return String.Empty;

            var sb = new StringBuilder();

            foreach (char @char in value)
            {
                if (reservedCharacters.IndexOf(@char) == -1)
                    sb.Append(@char);
                else
                    sb.AppendFormat("%{0:X2}", (int)@char);
            }
            return sb.ToString();
        }
    }
}

