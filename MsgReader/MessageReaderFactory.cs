using System.Web;

namespace DocumentServices.Modules.Readers.MsgReader
{
    public static class MessageReaderFactory
    {
        public static IReader CreateMessageReader()
        {
            return new Reader();
        }

        public static IReader CreateMessageReader(HttpContext context)
        {
            return new WebReader(context.Request.RawUrl,
                                 context.Request.PhysicalApplicationPath,
                                 context.Request.ApplicationPath);
        }
    }
}