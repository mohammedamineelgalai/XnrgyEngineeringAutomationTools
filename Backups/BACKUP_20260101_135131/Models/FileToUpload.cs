using System;

namespace VaultSDK_DirectUpload
{
    /// <summary>
    /// Représente un fichier à uploader avec son chemin local et son chemin Vault calculé
    /// </summary>
    public class FileToUpload
    {
        /// <summary>
        /// Chemin complet local du fichier (ex: C:\Vault\Engineering\Projects\10359\REF09\M03\fichier.dxf)
        /// </summary>
        public string LocalPath { get; set; }

        /// <summary>
        /// Chemin Vault calculé (ex: $/Engineering/Projects/10359/REF09/M03)
        /// </summary>
        public string VaultFolderPath { get; set; }

        /// <summary>
        /// Nom du fichier uniquement
        /// </summary>
        public string FileName => System.IO.Path.GetFileName(LocalPath);

        /// <summary>
        /// Taille du fichier formatée
        /// </summary>
        public string FileSize { get; set; }

        /// <summary>
        /// Affichage dans la liste
        /// </summary>
        public string DisplayText => $"{FileName} → {VaultFolderPath} ({FileSize})";

        public FileToUpload(string localPath, string vaultFolderPath, long fileSize)
        {
            LocalPath = localPath;
            VaultFolderPath = vaultFolderPath;
            FileSize = FormatSize(fileSize);
        }

        private string FormatSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len /= 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        public override string ToString() => DisplayText;
    }
}

