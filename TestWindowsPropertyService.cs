using System;
using System.IO;
using VaultAutomationTool.Services;

namespace VaultAutomationTool
{
    /// <summary>
    /// Test simple pour vérifier WindowsPropertyService sur un fichier Inventor
    /// </summary>
    public class TestWindowsPropertyService
    {
        public static void TestWithFile(string testFilePath)
        {
            Console.WriteLine("╔══════════════════════════════════════════════════════════════════╗");
            Console.WriteLine("║     TEST WindowsPropertyService - Propriétés OLE Windows        ║");
            Console.WriteLine("╚══════════════════════════════════════════════════════════════════╝");
            Console.WriteLine();

            if (string.IsNullOrEmpty(testFilePath))
            {
                Console.WriteLine("Usage: Fournir le chemin d'un fichier Inventor (.ipt, .iam, .idw, .ipn)");
                return;
            }

            if (!File.Exists(testFilePath))
            {
                Console.WriteLine($"❌ Fichier non trouvé: {testFilePath}");
                return;
            }

            Console.WriteLine($"📁 Fichier test: {testFilePath}");
            Console.WriteLine($"   Extension: {Path.GetExtension(testFilePath)}");
            Console.WriteLine($"   Taille: {new FileInfo(testFilePath).Length:N0} bytes");
            Console.WriteLine();

            // Valeurs de test
            string testProject = "TEST-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string testReference = "REF-WINPROP";
            string testModule = "M999";

            Console.WriteLine("📝 Valeurs à écrire:");
            Console.WriteLine($"   Project   = {testProject}");
            Console.WriteLine($"   Reference = {testReference}");
            Console.WriteLine($"   Module    = {testModule}");
            Console.WriteLine();

            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine("ÉCRITURE des propriétés...");
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

            try
            {
                using (var propService = new WindowsPropertyService())
                {
                    bool result = propService.SetCustomProperties(testFilePath, testProject, testReference, testModule);
                    
                    Console.WriteLine();
                    if (result)
                    {
                        Console.WriteLine("✅ Propriétés écrites avec SUCCÈS!");
                    }
                    else
                    {
                        Console.WriteLine("❌ Échec de l'écriture des propriétés");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Exception lors de l'écriture: {ex.Message}");
                Console.WriteLine($"   Stack: {ex.StackTrace}");
                return;
            }

            Console.WriteLine();
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine("LECTURE de vérification...");
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

            try
            {
                using (var propService = new WindowsPropertyService())
                {
                    string readProject = propService.ReadCustomProperty(testFilePath, "Project");
                    string readReference = propService.ReadCustomProperty(testFilePath, "Reference");
                    string readModule = propService.ReadCustomProperty(testFilePath, "Module");

                    Console.WriteLine($"   Project lu   = {readProject ?? "(null)"}");
                    Console.WriteLine($"   Reference lu = {readReference ?? "(null)"}");
                    Console.WriteLine($"   Module lu    = {readModule ?? "(null)"}");
                    Console.WriteLine();

                    bool projectOk = readProject == testProject;
                    bool referenceOk = readReference == testReference;
                    bool moduleOk = readModule == testModule;

                    Console.WriteLine("Vérification:");
                    Console.WriteLine($"   Project   : {(projectOk ? "✅" : "❌")}");
                    Console.WriteLine($"   Reference : {(referenceOk ? "✅" : "❌")}");
                    Console.WriteLine($"   Module    : {(moduleOk ? "✅" : "❌")}");
                    Console.WriteLine();

                    if (projectOk && referenceOk && moduleOk)
                    {
                        Console.WriteLine("✅✅✅ TEST RÉUSSI - Les propriétés Windows OLE fonctionnent! ✅✅✅");
                    }
                    else
                    {
                        Console.WriteLine("⚠️ TEST PARTIEL - Certaines propriétés n'ont pas été lues correctement");
                        Console.WriteLine("   Note: Le format Inventor 2026 peut utiliser un encodage propriétaire");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Exception lors de la lecture: {ex.Message}");
            }

            Console.WriteLine();
            Console.WriteLine("═════════════════════════════════════════════════════════════════════");
            Console.WriteLine("Test terminé - Vérifiez le fichier dans Inventor pour confirmer");
            Console.WriteLine("═════════════════════════════════════════════════════════════════════");
        }
    }
}
