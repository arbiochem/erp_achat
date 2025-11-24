namespace arbioApp.Utils.Global
{
    public class ConnectedUser
    {
        public static int UserId;
        public static string UserName;
        public static RoleUser roles;
    }
    public enum RoleUser
    {
        Admin,
        Caissier
    }
}
