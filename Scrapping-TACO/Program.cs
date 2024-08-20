// See https://aka.ms/new-console-template for more information
using ExcelDataReader;
using Microsoft.EntityFrameworkCore;
using Scrapping_TACO.Models;
using System.Globalization;

class Program {
    static async Task Main(string[] args) {
        Console.WriteLine("Hello, World!");
        var optionsBuilder = new DbContextOptionsBuilder<balancedlifeContext>();
        optionsBuilder.UseSqlServer("Data Source=localhost\\MSSQLSERVER02;Initial Catalog=Balanced_Life;Integrated Security=True;Encrypt=True;TrustServerCertificate=YES");


        string filePath = @"C:\Users\FMX\source\repos\Scrapping-TACO\Scrapping-TACO\Taco_4a_edicao_2011 (2).xls";

        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var nutrientMappings = new Dictionary<int, (int NutrientId, int UnitId)>
{
            { 1, (39, 10004) }, // Umidade
            { 2, (38, 10003) }, // Energia (kcal)
            { 3, (38, 10002) }, // Energia (kJ)
            { 4, (42, 10004) }, // Proteína
            { 5, (43, 10004) }, // Lipídeos
            { 6, (47, 10005) }, // Colesterol
            { 7, (40, 10004) }, // Carboidrato
            { 8, (44, 10004) }, // Fibra alimentar
            { 9, (46, 10004) }, // Cinzas
            { 10, (52, 10005) }, // Cálcio
            { 11, (55, 10005) }, // Magnésio
            { 12, (58, 10005) }, // Manganês
            { 13, (56, 10005) }, // Fósforo
            { 14, (53, 10005) }, // Ferro
            { 15, (54, 10005) }, // Sódio
            { 16, (57, 10005) }, // Potássio
            { 17, (60, 10005) }, // Cobre
            { 18, (59, 10005) }, // Zinco
            { 19, (62, 10008) }, // Retinol
            { 20, (63, 10008) }, // RE
            { 21, (64, 10008) }, // RAE
            { 22, (66, 10005) }, // Tiamina
            { 23, (67, 10005) }, // Riboflavina
            { 24, (69, 10005) }, // Piridoxina
            { 25, (68, 10005) }, // Niacina
            { 26, (71, 10005) }  // Vitamina C
        };

        using ( var stream = File.Open(filePath, FileMode.Open, FileAccess.Read) ) {
            // Cria o leitor de Excel
            using ( var reader = ExcelReaderFactory.CreateReader(stream) )

            using ( var dbContext = new balancedlifeContext(optionsBuilder.Options) ) {
                
                while ( reader.Read() ) {
                    // Lê os dados das colunas
                    var nomeAlimento = reader.GetValue(0)?.ToString();
                    if (string.IsNullOrEmpty(nomeAlimento)) 
                        continue;
                    
                    var umidade = reader.GetValue(1).ToSafeDouble();
                    var energiaKcal = reader.GetValue(2).ToSafeDouble();
                    var energiaKj = reader.GetValue(3).ToSafeDouble();
                    var proteina = reader.GetValue(4).ToSafeDouble();
                    var lipideos = reader.GetValue(5).ToSafeDouble();
                    var colesterol = reader.GetValue(6).ToSafeDouble();
                    var carboidrato = reader.GetValue(7).ToSafeDouble();
                    var fibraAlimentar = reader.GetValue(8).ToSafeDouble();
                    var cinzas = reader.GetValue(9).ToSafeDouble();
                    var calcio = reader.GetValue(10).ToSafeDouble();
                    var magnesio = reader.GetValue(11).ToSafeDouble();
                    var manganesio = reader.GetValue(12).ToSafeDouble();
                    var fosforo = reader.GetValue(13).ToSafeDouble();
                    var ferro = reader.GetValue(14).ToSafeDouble();
                    var sodio = reader.GetValue(15).ToSafeDouble();
                    var potassio = reader.GetValue(16).ToSafeDouble();
                    var cobre = reader.GetValue(17).ToSafeDouble();
                    var zinco = reader.GetValue(18).ToSafeDouble();
                    var retinol = reader.GetValue(19).ToSafeDouble();
                    var re = reader.GetValue(20).ToSafeDouble();
                    var rae = reader.GetValue(21).ToSafeDouble();
                    var tiamina = reader.GetValue(22).ToSafeDouble();
                    var riboflavina = reader.GetValue(23).ToSafeDouble();
                    var piridoxina = reader.GetValue(24).ToSafeDouble();
                    var niacina = reader.GetValue(25).ToSafeDouble();
                    var vitaminaC = reader.GetValue(26).ToSafeDouble();
                    var foodGroupId = reader.GetValue(27).ToSafeInt();

                    Console.WriteLine(umidade);
                    Console.WriteLine(energiaKcal);
                    Console.WriteLine(energiaKj);
                    Console.WriteLine(proteina);
                    Console.WriteLine(lipideos);
                    Console.WriteLine(colesterol);
                    Console.WriteLine(carboidrato);
                    Console.WriteLine(fibraAlimentar);
                    Console.WriteLine(cinzas);
                    Console.WriteLine(calcio);
                    Console.WriteLine(magnesio);
                    Console.WriteLine(manganesio);
                    Console.WriteLine(fosforo);
                    Console.WriteLine(ferro);
                    Console.WriteLine(sodio);
                    Console.WriteLine(potassio);
                    Console.WriteLine(cobre);
                    Console.WriteLine(zinco);
                    Console.WriteLine(retinol);
                    Console.WriteLine(re);
                    Console.WriteLine(rae);
                    Console.WriteLine(tiamina);
                    Console.WriteLine(riboflavina);
                    Console.WriteLine(piridoxina);
                    Console.WriteLine(niacina);
                    Console.WriteLine(vitaminaC);
                    Console.WriteLine(foodGroupId);
                    Console.WriteLine("--------------------------------------------------------------------------------------------------");

                    // Criar o objeto Food
                    var food = new Food {
                        Name = nomeAlimento,
                        Brand = "",
                        IdFoodGroup = foodGroupId,
                        ReferenceTable = "TACO",
                    };

                    // Adicionar ao contexto
                    dbContext.Foods.Add(food);
                    Console.WriteLine(food.Name);

                    for ( int i = 1; i <= 26; i++ ) {
                        if ( nutrientMappings.TryGetValue(i, out var mapping) ) {
                            var value = reader.GetValue(i)?.ToSafeDouble() ?? 0;
                            if ( i == 3 ) 
                                continue;
                            
                            var nutritionInfo = new FoodNutritionInfo {
                                IdFood = food.Id,
                                IdFoodNavigation = food,
                                IdNutritionalComposition = mapping.NutrientId,
                                Quantity = value,
                                IdUnitMeasurement = mapping.UnitId
                            };

                            dbContext.FoodNutritionInfos.Add(nutritionInfo);
                        }
                    }
                }
                dbContext.SaveChanges();
            }
        }
    }
}