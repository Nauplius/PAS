using Microsoft.Office.Server.PowerPoint.Conversion;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System;
using System.ServiceModel;

namespace Nauplius.PAS
{
    internal class Conversion
    {
        public static bool ConvertToFormat(string siteUrl, SPFile file, SPFileCollection fileCollection,
            string sourceExtension, PresentationType pType, SPFolder folder, string outFileName,
            bool wait)
        {
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        if (file != null)
                        {
                            var fStream = file.OpenBinaryStream();
                            var sStream = new SPFileStream(web, 0x1000);
                            var request = new PresentationRequest(fStream, sourceExtension, pType, sStream);
                            var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                            request.EndConvert(result);

                            if (web.Url != folder.ParentWeb.Url)
                            {
                                using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }
                            }
                            else
                            {
                                folder.Files.Add(outFileName, sStream, true);
                            }

                            return true;
                        }

                        if (fileCollection != null)
                        {
                            foreach (SPFile file1 in fileCollection)
                            {
                                var fStream = file1.OpenBinaryStream();
                                var sStream = new SPFileStream(web, 0x1000);
                                var request = new PresentationRequest(fStream, sourceExtension, pType, sStream);
                                var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                                request.EndConvert(result);

                                if (web.Url != folder.ParentWeb.Url)
                                {
                                    using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                    {
                                        folder.Files.Add(outFileName, sStream, true);
                                    }
                                }
                                else
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }

                                return true;
                            }
                        }
                    }
                    catch (SPException exception)
                    {
                        Exceptions.SharePointException(exception, file);
                        return false;
                    }
                    catch (ServiceActivationException exception)
                    {
                        Exceptions.ServiceException(exception, file);
                        return false;
                    }
                    catch (CommunicationException exception)
                    {
                        Exceptions.CommException(exception, file);
                        return false;
                    }
                    catch (ConversionException exception)
                    {
                        Exceptions.ConversionException(exception, file);
                        return false;
                    }
                    catch (Exception exception)
                    {
                        Exceptions.GenericException(exception, file);
                        return false;
                    }
                }
            }

            return false;
        }

        public static bool ConvertToPicture(string siteUrl, SPFile file, SPFileCollection fileCollection,
            string sourceExtension, PictureFormat pictureFormat, SPFolder folder, string outFileName)
        {
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        if (file != null)
                        {
                            var fStream = file.OpenBinaryStream();
                            var sStream = new SPFileStream(web, 0x1000);
                            var request = new PictureRequest(fStream, sourceExtension, pictureFormat, sStream);
                            var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                            request.EndConvert(result);

                            if (web.Url != folder.ParentWeb.Url)
                            {
                                using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }
                            }
                            else
                            {
                                folder.Files.Add(outFileName, sStream, true);
                            }

                            return true;
                        }

                        if (fileCollection != null)
                        {
                            foreach (SPFile file1 in fileCollection)
                            {
                                var fStream = file1.OpenBinaryStream();
                                var sStream = new SPFileStream(web, 0x1000);
                                var request = new PictureRequest(fStream, sourceExtension, pictureFormat, sStream);
                                var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                                request.EndConvert(result);

                                if (web.Url != folder.ParentWeb.Url)
                                {
                                    using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                    {
                                        folder.Files.Add(outFileName, sStream, true);
                                    }
                                }
                                else
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }

                                return true;
                            }
                        }
                    }
                    catch (SPException exception)
                    {
                        Exceptions.SharePointException(exception, file);
                        return false;
                    }
                    catch (ServiceActivationException exception)
                    {
                        Exceptions.ServiceException(exception, file);
                        return false;
                    }
                    catch (CommunicationException exception)
                    {
                        Exceptions.CommException(exception, file);
                        return false;
                    }
                    catch (Exception exception)
                    {
                        Exceptions.GenericException(exception, file);
                        return false;
                    }
                }
            }
            return false;
        }

        public static bool ConvertToPdf(string siteUrl, SPFile file, SPFileCollection fileCollection,
            string sourceExtension, SPFolder folder, FixedFormatSettings fixedFormatSettings, string outFileName)
        {
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        if (file != null)
                        {
                            var fStream = file.OpenBinaryStream();
                            var sStream = new SPFileStream(web, 0x1000);
                            var request = new PdfRequest(fStream, sourceExtension, fixedFormatSettings, sStream);
                            var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                            request.EndConvert(result);

                            if (web.Url != folder.ParentWeb.Url)
                            {
                                using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }
                            }
                            else
                            {
                                folder.Files.Add(outFileName, sStream, true);
                            }

                            return true;
                        }

                        if (fileCollection != null)
                        {
                            foreach (SPFile file1 in fileCollection)
                            {
                                var fStream = file1.OpenBinaryStream();
                                var sStream = new SPFileStream(web, 0x1000);
                                var request = new PdfRequest(fStream, sourceExtension, fixedFormatSettings, sStream);
                                var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                                request.EndConvert(result);

                                if (web.Url != folder.ParentWeb.Url)
                                {
                                    using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                    {
                                        folder.Files.Add(outFileName, sStream, true);
                                    }
                                }
                                else
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }

                                return true;
                            }
                        }
                    }
                    catch (SPException exception)
                    {
                        Exceptions.SharePointException(exception, file);
                        return false;
                    }
                    catch (ServiceActivationException exception)
                    {
                        Exceptions.ServiceException(exception, file);
                        return false;
                    }
                    catch (CommunicationException exception)
                    {
                        Exceptions.CommException(exception, file);
                        return false;
                    }
                    catch (Exception exception)
                    {
                        Exceptions.GenericException(exception, file);
                        return false;
                    }
                }
            }
            return false;
        }

        public static bool ConverToXps(string siteUrl, SPFile file, SPFileCollection fileCollection,
            string sourceExtension, SPFolder folder, FixedFormatSettings fixedFormatSettings, string outFileName)
        {
            using (SPSite site = new SPSite(siteUrl))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        if (file != null)
                        {
                            var fStream = file.OpenBinaryStream();
                            var sStream = new SPFileStream(web, 0x1000);
                            var request = new XpsRequest(fStream, sourceExtension, fixedFormatSettings, sStream);
                            var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                            request.EndConvert(result);

                            if (web.Url != folder.ParentWeb.Url)
                            {
                                using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }
                            }
                            else
                            {
                                folder.Files.Add(outFileName, sStream, true);
                            }

                            return true;
                        }

                        if (fileCollection != null)
                        {
                            foreach (SPFile file1 in fileCollection)
                            {
                                var fStream = file1.OpenBinaryStream();
                                var sStream = new SPFileStream(web, 0x1000);
                                var request = new XpsRequest(fStream, sourceExtension, fixedFormatSettings, sStream);
                                var result = request.BeginConvert(SPServiceContext.GetContext(site), null, null);
                                request.EndConvert(result);

                                if (web.Url != folder.ParentWeb.Url)
                                {
                                    using (SPWeb web2 = site.OpenWeb(folder.ParentWeb.Url))
                                    {
                                        folder.Files.Add(outFileName, sStream, true);
                                    }
                                }
                                else
                                {
                                    folder.Files.Add(outFileName, sStream, true);
                                }

                                return true;
                            }
                        }
                    }
                    catch (SPException exception)
                    {
                        Exceptions.SharePointException(exception, file);
                        return false;
                    }
                    catch (ServiceActivationException exception)
                    {
                        Exceptions.ServiceException(exception, file);
                        return false;
                    }
                    catch (CommunicationException exception)
                    {
                        Exceptions.CommException(exception, file);
                        return false;
                    }
                    catch (ConversionException exception)
                    {
                        Exceptions.ConversionException(exception, file);
                        return false;
                    }
                    catch (Exception exception)
                    {
                        Exceptions.GenericException(exception, file);
                        return false;
                    }
                }
            }

            return false;
        }
}

    class Exceptions
    {
        internal static void SharePointException(SPException exception, SPFile file)
        {
            SPDiagnosticsService.Local.WriteTrace(0, new SPDiagnosticsCategory("NaupliusPASStatus",
                TraceSeverity.High, EventSeverity.Error), TraceSeverity.Unexpected,
                "An unexpected error has occurred. " + exception.StackTrace);
            SPUtility.TransferToErrorPage(string.Format("An error occurred while converting {0}. Please contact your SharePoint Administrator.", file.Name));
        }

        internal static void ServiceException(ServiceActivationException exception, SPFile file)
        {
            SPDiagnosticsService.Local.WriteTrace(0, new SPDiagnosticsCategory("NaupliusPASStatus",
                TraceSeverity.High, EventSeverity.Error), TraceSeverity.Unexpected,
                "An unexpected error has occurred with the PowerPoint Service Application. " +
                "Check the status of the PowerPoint Service Instance and Service Application. " + exception.StackTrace);
            SPUtility.TransferToErrorPage(string.Format("An error occurred while converting {0}. Please contact your SharePoint Administrator.", file.Name));
        }

        internal static void ConversionException(ConversionException exception, SPFile file)
        {
            SPDiagnosticsService.Local.WriteTrace(0, new SPDiagnosticsCategory("NaupliusPASStatus",
                TraceSeverity.High, EventSeverity.Error), TraceSeverity.Unexpected,
                "An unexpected error has occurred with the conversion. " +
                exception.Message + exception.StackTrace);
            
            SPUtility.TransferToErrorPage(string.Format("An error occurred while coverting {0}. Please contact your SharePoint Administrator.", file.Name));
        }
        internal static void CommException(CommunicationException exception, SPFile file)
        {
            SPDiagnosticsService.Local.WriteTrace(0, new SPDiagnosticsCategory("NaupliusPASStatus",
                TraceSeverity.High, EventSeverity.Error), TraceSeverity.Unexpected,
                "An unexpected error has occurred attempting to communicate with the PowerPoint Web Service. " +
                "Check the status of the PowerPoint Service Instance and hosting IIS Application Pool. " +  exception.StackTrace);
            SPUtility.TransferToErrorPage(string.Format("An error occurred while converting {0}. Please contact your SharePoint Administrator.", file.Name));
        }

        internal static void GenericException(Exception exception, SPFile file)
        {
            SPDiagnosticsService.Local.WriteTrace(0, new SPDiagnosticsCategory("NaupliusPASStatus",
                TraceSeverity.High, EventSeverity.Error), TraceSeverity.Unexpected, exception.Message + " " + exception.StackTrace);
            SPUtility.TransferToErrorPage(string.Format("An error occurred while converting {0}. Please contact your SharePoint Administrator.", file.Name));
        }
    }
}
