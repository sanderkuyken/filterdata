// Wacht tot het document geladen is
document.addEventListener('DOMContentLoaded', function() {
    // Functie om leeftijdsselecties te genereren
    function generateAgeSelections() {
        const selections = [
            { label: "Alle leeftijden", min: 0, max: 150 },
            { label: "0-12 jaar", min: 0, max: 12 },
            { label: "13-17 jaar", min: 13, max: 17 },
            { label: "18-25 jaar", min: 18, max: 25 },
            { label: "26-35 jaar", min: 26, max: 35 },
            { label: "36-50 jaar", min: 36, max: 50 },
            { label: "51-65 jaar", min: 51, max: 65 },
            { label: "66+ jaar", min: 66, max: 150 },
        ];

        return selections.map(selection =>
            `<option value="${selection.min},${selection.max}">${selection.label}</option>`
        ).join('');
    }

    // Functie om leeftijd te berekenen
    function calculateAge(birthDate) {
        const today = new Date();
        const birth = new Date(birthDate);

        const birthDateConverted = (typeof birthDate === 'number')
            ? new Date(Math.round((birthDate - 25569) * 86400 * 1000))
            : birth;

        let age = today.getFullYear() - birthDateConverted.getFullYear();
        const monthDiff = today.getMonth() - birthDateConverted.getMonth();

        if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDateConverted.getDate())) {
            age--;
        }

        return age;
    }

    // Functie om geslacht statistieken te berekenen
    function calculateGenderStats(data) {
        const total = data.length;
        const genderCount = {
            male: 0,
            female: 0,
            unknown: 0
        };

        data.forEach(row => {
            const gender = (row.Geslacht || row.geslacht || row.Gender || row.gender || '').toLowerCase().trim();
            if (gender === 'm' || gender === 'man' || gender === 'male' || gender === 'h') {
                genderCount.male++;
            } else if (gender === 'v' || gender === 'vrouw' || gender === 'female' || gender === 'f') {
                genderCount.female++;
            } else {
                genderCount.unknown++;
            }
        });

        return {
            total,
            male: {
                count: genderCount.male,
                percentage: ((genderCount.male / total) * 100).toFixed(1)
            },
            female: {
                count: genderCount.female,
                percentage: ((genderCount.female / total) * 100).toFixed(1)
            },
            unknown: {
                count: genderCount.unknown,
                percentage: ((genderCount.unknown / total) * 100).toFixed(1)
            }
        };
    }

    // Functie om leeftijdsstatistieken te berekenen
    function calculateAgeStats(data) {
        const totalCount = data.length;
        const ageGroups = {};

        data.forEach(row => {
            const age = row.Leeftijd;
            ageGroups[age] = (ageGroups[age] || 0) + 1;
        });

        const stats = Object.entries(ageGroups).map(([age, count]) => ({
            age: parseInt(age),
            count: count,
            percentage: ((count / totalCount) * 100).toFixed(1)
        }));

        return stats.sort((a, b) => a.age - b.age);
    }

    // Functie voor woonplaats statistieken
    // Update de calculateCityStats functie met betere type checking
    function calculateCityStats(data) {
        const total = data.length;
        const cityCount = {};

        data.forEach(row => {
            // Converteer eerst naar string en handel null/undefined af
            let city = (row.Woonplaats || row.woonplaats || row.Stad || row.stad || row.City || row.city);

            // Zorg ervoor dat we een string hebben en trim deze
            city = String(city || '').trim();

            if (city) {
                cityCount[city] = (cityCount[city] || 0) + 1;
            } else {
                cityCount['Onbekend'] = (cityCount['Onbekend'] || 0) + 1;
            }
        });

        return Object.entries(cityCount)
            .map(([city, count]) => ({
                city,
                count,
                percentage: ((count / total) * 100).toFixed(1)
            }))
            .sort((a, b) => b.count - a.count);
    }

    // Functie om statistieken tabel te maken
    function createStatsTable(stats, sheetName, genderStats, cityStats) {
        return `
            <div class="mb-4">
                <h3 class="mb-3">Statistieken voor werkblad: ${sheetName}</h3>
                
                <!-- Gender statistieken -->
                <div class="alert alert-info mb-3">
                    <h4>Geslacht verdeling:</h4>
                    Mannen: ${genderStats.male.count} (${genderStats.male.percentage}%)<br>
                    Vrouwen: ${genderStats.female.count} (${genderStats.female.percentage}%)<br>
                    ${genderStats.unknown.count > 0 ? `Onbekend: ${genderStats.unknown.count} (${genderStats.unknown.percentage}%)<br>` : ''}
                </div>

                <!-- Woonplaats statistieken -->
                <div class="mb-4">
                    <h4>Woonplaats verdeling:</h4>
                    <table class="table table-striped">
                        <thead>
                            <tr>
                                <th>Woonplaats</th>
                                <th>Aantal</th>
                                <th>Percentage</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${cityStats.map(stat => `
                                <tr>
                                    <td>${stat.city}</td>
                                    <td>${stat.count}</td>
                                    <td>${stat.percentage}%</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
                
                <!-- Leeftijd statistieken -->
                <h4>Leeftijdsverdeling:</h4>
                <table class="table table-striped">
                    <thead>
                        <tr>
                            <th>Leeftijd</th>
                            <th>Aantal</th>
                            <th>Percentage</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${stats.map(stat => `
                            <tr>
                                <td>${stat.age} jaar</td>
                                <td>${stat.count}</td>
                                <td>${stat.percentage}%</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        `;
    }

    // Functie om Excel/CSV bestand te verwerken
    function processExcelFile(file, minAge, maxAge) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                const results = {};
                const newWorkbook = XLSX.utils.book_new();
                const totalCityCounts = {};

                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);

                    const processedData = jsonData.map(row => {
                        const birthDate = row.Geboortedatum || row.geboortedatum || row['Geboorte datum'] || row['geboorte datum'];
                        const age = calculateAge(birthDate);
                        return {
                            ...row,
                            Leeftijd: age
                        };
                    });

                    const filteredData = processedData.filter(row => {
                        return row.Leeftijd >= minAge && row.Leeftijd <= maxAge;
                    });

                    const stats = calculateAgeStats(filteredData);
                    const genderStats = calculateGenderStats(filteredData);
                    const cityStats = calculateCityStats(filteredData);

                    cityStats.forEach(stat => {
                        totalCityCounts[stat.city] = (totalCityCounts[stat.city] || 0) + stat.count;
                    });

                    const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
                    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

                    results[sheetName] = {
                        stats,
                        genderStats,
                        cityStats,
                        totalOriginal: jsonData.length,
                        totalFiltered: filteredData.length
                    };
                });

                const totalFilteredCount = Object.values(results).reduce((sum, result) => sum + result.totalFiltered, 0);
                const totalCityStats = Object.entries(totalCityCounts)
                    .map(([city, count]) => ({
                        city,
                        count,
                        percentage: ((count / totalFilteredCount) * 100).toFixed(1)
                    }))
                    .sort((a, b) => b.count - a.count);

                resolve({
                    excelBuffer: XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' }),
                    results,
                    totalCityStats
                });
            };

            reader.onerror = function(e) {
                reject(new Error('Fout bij het lezen van het bestand'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    // HTML voor de gebruikersinterface
    document.body.innerHTML = `
        <div class="container p-4">
            <h2>Excel Filter met Statistieken (2024)</h2>
            <div class="mb-3">
                <label for="excelFile">Selecteer Excel bestand:</label>
                <input type="file" id="excelFile" accept=".xlsx, .xls, .csv" class="form-control">
            </div>
            
            <div class="mb-3">
                <label for="ageRange">Selecteer leeftijdscategorie:</label>
                <select id="ageRange" class="form-control">
                    ${generateAgeSelections()}
                </select>
            </div>
            
            <div class="mb-3">
                <label for="customRange">Of voer aangepaste leeftijdsrange in:</label>
                <div class="row">
                    <div class="col">
                        <input type="number" id="minAge" class="form-control" placeholder="Min leeftijd" value="0">
                    </div>
                    <div class="col">
                        <input type="number" id="maxAge" class="form-control" placeholder="Max leeftijd" value="100">
                    </div>
                </div>
            </div>
            
            <button id="filterButton" class="btn btn-primary">Filter Toepassen</button>
            
            <div id="allStats" class="mt-4">
                <div id="totalSummary" class="alert alert-info" style="display: none;"></div>
                <div id="sheetStats"></div>
            </div>
        </div>
    `;

    // Event listeners
    document.getElementById('ageRange').addEventListener('change', function() {
        const [min, max] = this.value.split(',');
        document.getElementById('minAge').value = min;
        document.getElementById('maxAge').value = max;
    });

    document.getElementById('filterButton').addEventListener('click', async () => {
        const fileInput = document.getElementById('excelFile');
        const minAge = parseInt(document.getElementById('minAge').value);
        const maxAge = parseInt(document.getElementById('maxAge').value);

        if (!fileInput?.files?.length) {
            alert('Selecteer eerst een bestand');
            return;
        }

        try {
            const result = await processExcelFile(fileInput.files[0], minAge, maxAge);

            let totalOriginal = 0;
            let totalFiltered = 0;
            const totalGenderStats = {
                male: 0,
                female: 0,
                unknown: 0
            };

            Object.values(result.results).forEach(sheetResult => {
                totalOriginal += sheetResult.totalOriginal;
                totalFiltered += sheetResult.totalFiltered;
                totalGenderStats.male += sheetResult.genderStats.male.count;
                totalGenderStats.female += sheetResult.genderStats.female.count;
                totalGenderStats.unknown += sheetResult.genderStats.unknown.count;
            });

            const totalGenderPercentages = {
                male: ((totalGenderStats.male / totalFiltered) * 100).toFixed(1),
                female: ((totalGenderStats.female / totalFiltered) * 100).toFixed(1),
                unknown: ((totalGenderStats.unknown / totalFiltered) * 100).toFixed(1)
            };

            // Update totaaloverzicht
            const totalSummary = document.getElementById('totalSummary');
            totalSummary.style.display = 'block';
            totalSummary.innerHTML = `
                <h4>Totaal overzicht alle werkbladen</h4>
                Totaal aantal records: ${totalOriginal}<br>
                Aantal na filtering: ${totalFiltered}<br>
                Percentage geselecteerd: ${((totalFiltered / totalOriginal) * 100).toFixed(1)}%<br><br>
                <strong>Totale geslacht verdeling:</strong><br>
                Mannen: ${totalGenderStats.male} (${totalGenderPercentages.male}%)<br>
                Vrouwen: ${totalGenderStats.female} (${totalGenderPercentages.female}%)<br>
                ${totalGenderStats.unknown > 0 ? `Onbekend: ${totalGenderStats.unknown} (${totalGenderPercentages.unknown}%)<br>` : ''}<br>
                <strong>Totale woonplaats verdeling:</strong>
                <table class="table table-striped mt-2">
                    <thead>
                        <tr>
                            <th>Woonplaats</th>
                            <th>Aantal</th>
                            <th>Percentage</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${result.totalCityStats.map(stat => `
                            <tr>
                                <td>${stat.city}</td>
                                <td>${stat.count}</td>
                                <td>${stat.percentage}%</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            `;

            // Update werkblad statistieken
            const sheetStats = document.getElementById('sheetStats');
            sheetStats.innerHTML = '';
            Object.entries(result.results).forEach(([sheetName, sheetResult]) => {
                const sheetSummary = `
                    <div class="alert alert-secondary mb-3">
                        Werkblad: ${sheetName}<br>
                        Aantal records: ${sheetResult.totalOriginal}<br>
                        Na filtering: ${sheetResult.totalFiltered}<br>
                        Percentage: ${((sheetResult.totalFiltered / sheetResult.totalOriginal) * 100).toFixed(1)}%
                    </div>
                `;
                sheetStats.innerHTML += sheetSummary + createStatsTable(
                    sheetResult.stats,
                    sheetName,
                    sheetResult.genderStats,
                    sheetResult.cityStats
                );
            });

            // Download het gefilterde bestand
            const blob = new Blob([result.excelBuffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'gefilterde_data_2024.xlsx';
            a.click();
            URL.revokeObjectURL(url);

        } catch (error) {
            alert('Er is een fout opgetreden: ' + error.message);
            console.error('Error:', error);
        }
    });
});
// Functie voor het genereren van leeftijdsgroep selecties
function addAgeGroupSelection() {
    return `
        <div class="age-group-row mb-2">
            <div class="row align-items-center">
                <div class="col">
                    <input type="text" class="form-control group-name" placeholder="Naam voor tabblad" required>
                </div>
                <div class="col">
                    <input type="number" class="form-control min-age" placeholder="Min leeftijd" required>
                </div>
                <div class="col">
                    <input type="number" class="form-control max-age" placeholder="Max leeftijd" required>
                </div>
                <div class="col-auto">
                    <button type="button" class="btn btn-danger remove-group">×</button>
                </div>
            </div>
        </div>
    `;
}

// Update de HTML interface
document.body.innerHTML = `
    <div class="container p-4">
        <h2>Excel Filter met Statistieken (2024)</h2>
        
        <!-- Bestand selectie -->
        <div class="mb-3">
            <label for="excelFile">Selecteer Excel bestand:</label>
            <input type="file" id="excelFile" accept=".xlsx, .xls, .csv" class="form-control">
        </div>
        
        <!-- Standaard leeftijdsfilter -->
        <div class="mb-3">
            <label for="ageRange">Selecteer leeftijdscategorie voor hoofdfilter:</label>
            <select id="ageRange" class="form-control">
                ${generateAgeSelections()}
            </select>
        </div>
        
        <!-- Aangepaste leeftijdsgroepen -->
        <div class="card mb-3">
            <div class="card-header">
                <h5 class="mb-0">Aangepaste leeftijdsgroepen</h5>
                <small class="text-muted">Deze worden als aparte tabbladen toegevoegd</small>
            </div>
            <div class="card-body">
                <div id="ageGroups">
                    <!-- Hier komen de leeftijdsgroep rijen -->
                </div>
                <button type="button" id="addAgeGroup" class="btn btn-secondary">
                    + Leeftijdsgroep toevoegen
                </button>
            </div>
        </div>
        
        <button id="filterButton" class="btn btn-primary">Filter Toepassen</button>
        
        <div id="allStats" class="mt-4">
            <div id="totalSummary" class="alert alert-info" style="display: none;"></div>
            <div id="sheetStats"></div>
        </div>
    </div>
`;


// Update de processExcelFile functie
function processExcelFile(file, minAge, maxAge, customAgeGroups) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const results = {};
            const newWorkbook = XLSX.utils.book_new();
            const totalCityCounts = {};

            // Verwerk elk origineel werkblad
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                // Basisverwerking voor hoofdfilter
                const processedData = jsonData.map(row => {
                    const birthDate = row.Geboortedatum || row.geboortedatum || row['Geboorte datum'] || row['geboorte datum'];
                    const age = calculateAge(birthDate);
                    return { ...row, Leeftijd: age };
                });

                const filteredData = processedData.filter(row =>
                    row.Leeftijd >= minAge && row.Leeftijd <= maxAge
                );

                // Voeg basis gefilterde data toe
                const newWorksheet = XLSX.utils.json_to_sheet(filteredData);
                XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

                // Bereken statistieken voor hoofdfilter
                const stats = calculateAgeStats(filteredData);
                const genderStats = calculateGenderStats(filteredData);
                const cityStats = calculateCityStats(filteredData);

                results[sheetName] = {
                    stats,
                    genderStats,
                    cityStats,
                    totalOriginal: jsonData.length,
                    totalFiltered: filteredData.length
                };

                // Verwerk aangepaste leeftijdsgroepen
                customAgeGroups.forEach(group => {
                    const groupFilteredData = processedData.filter(row =>
                        row.Leeftijd >= group.minAge && row.Leeftijd <= group.maxAge
                    );

                    const groupStats = {
                        ageStats: calculateAgeStats(groupFilteredData),
                        genderStats: calculateGenderStats(groupFilteredData),
                        cityStats: calculateCityStats(groupFilteredData),
                        totalFiltered: groupFilteredData.length,
                        percentage: ((groupFilteredData.length / jsonData.length) * 100).toFixed(1)
                    };

                    // Voeg werkblad toe voor deze leeftijdsgroep
                    const groupWorksheet = XLSX.utils.json_to_sheet(groupFilteredData);
                    const groupSheetName = `${sheetName}_${group.name}`;
                    XLSX.utils.book_append_sheet(newWorkbook, groupWorksheet, groupSheetName);

                    // Sla statistieken op
                    results[groupSheetName] = {
                        ...groupStats,
                        totalOriginal: jsonData.length,
                        isCustomGroup: true,
                        groupName: group.name,
                        ageRange: `${group.minAge}-${group.maxAge}`
                    };
                });
            });

            resolve({
                excelBuffer: XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' }),
                results: results
            });
        };

        reader.onerror = function(e) {
            reject(new Error('Fout bij het lezen van het bestand'));
        };

        reader.readAsArrayBuffer(file);
    });
}

// Update de event listeners
document.addEventListener('DOMContentLoaded', function() {
    // Add group button event listener
    document.getElementById('addAgeGroup').addEventListener('click', function() {
        const ageGroupsDiv = document.getElementById('ageGroups');
        const newGroupDiv = document.createElement('div');
        newGroupDiv.innerHTML = addAgeGroupSelection();
        ageGroupsDiv.appendChild(newGroupDiv);

        // Add remove button listener
        newGroupDiv.querySelector('.remove-group').addEventListener('click', function() {
            newGroupDiv.remove();
        });
    });

    // Filter button event listener
    document.getElementById('filterButton').addEventListener('click', async () => {
        const fileInput = document.getElementById('excelFile');
        const minAge = parseInt(document.getElementById('minAge').value);
        const maxAge = parseInt(document.getElementById('maxAge').value);

        if (!fileInput?.files?.length) {
            alert('Selecteer eerst een bestand');
            return;
        }

        // Verzamel alle aangepaste leeftijdsgroepen
        const customAgeGroups = [];
        document.querySelectorAll('.age-group-row').forEach(row => {
            const name = row.querySelector('.group-name').value.trim();
            const minAge = parseInt(row.querySelector('.min-age').value);
            const maxAge = parseInt(row.querySelector('.max-age').value);

            if (name && !isNaN(minAge) && !isNaN(maxAge)) {
                customAgeGroups.push({ name, minAge, maxAge });
            }
        });

        try {
            const result = await processExcelFile(fileInput.files[0], minAge, maxAge, customAgeGroups);

            // Update de statistieken weergave
            const sheetStats = document.getElementById('sheetStats');
            sheetStats.innerHTML = '';

            Object.entries(result.results).forEach(([sheetName, sheetResult]) => {
                const isCustomGroup = sheetResult.isCustomGroup;
                const sheetSummary = `
                    <div class="alert ${isCustomGroup ? 'alert-info' : 'alert-secondary'} mb-3">
                        <h4>${isCustomGroup ? `Leeftijdsgroep: ${sheetResult.groupName} (${sheetResult.ageRange} jaar)` : `Werkblad: ${sheetName}`}</h4>
                        Aantal records: ${sheetResult.totalOriginal}<br>
                        Na filtering: ${sheetResult.totalFiltered}<br>
                        Percentage: ${sheetResult.percentage || ((sheetResult.totalFiltered / sheetResult.totalOriginal) * 100).toFixed(1)}%
                    </div>
                `;
                sheetStats.innerHTML += sheetSummary + createStatsTable(
                    sheetResult.ageStats || sheetResult.stats,
                    sheetName,
                    sheetResult.genderStats,
                    sheetResult.cityStats
                );
            });

            // Download het gefilterde bestand
            const blob = new Blob([result.excelBuffer], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'gefilterde_data_2024.xlsx';
            a.click();
            URL.revokeObjectURL(url);

        } catch (error) {
            alert('Er is een fout opgetreden: ' + error.message);
            console.error('Error:', error);
        }
    });
});