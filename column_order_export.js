/**
 * Ensures columns are exported in the correct order
 */
function exportWithColumnOrder() {
    // Column name mapping for display
    const columnDisplayNames = {
        'commune': 'Commune',
        'region': 'Région',
        'totalParcelles': 'Total Parcelles',
        'percentTotal': '% du Total',
        'nicad': 'NICAD',
        'percentNicad': '% NICAD',
        'ctasf': 'CTASF',
        'percentCtasf': '% CTASF',
        'deliberated': 'Délibérées',
        'percentDeliberated': '% Délibérée',
        'parcellesBrutes': 'Parcelles brutes',
        'collected': 'Parcelles collectées (sans doublon géométrique)',
        'surveyed': 'Parcelles enquêtées',
        'rejectionReasons': 'Motifs de rejet post-traitement',
        'retained': 'Parcelles retenues après post-traitement',
        'validated': 'Parcelles validées par l\'URM',
        'rejected': 'Parcelles rejetées par l\'URM' ,
        'urmRejectionReasons': 'Motifs de rejet URM',
        'corrected': 'Parcelles corrigées',
        'geomaticien': 'Geomaticien',
        'individualJoined': 'Parcelles individuelles jointes',
        'collectiveJoined': 'Parcelles collectives jointes',
        'unjoined': 'Parcelles non jointes',
        'duplicatesRemoved': 'Doublons supprimés',
        'duplicateRemovalRate': 'Taux suppression doublons (%)',
        'parcelsInConflict': 'Parcelles en conflit',
        'significantDuplicates': 'Significant Duplicates',
        'postProcessedLot1_46': 'Parc. post-traitées lot 1-46',
        'jointureStatus': 'Statut jointure',
        'jointureErrorMessage': 'Message d\'erreur jointure'
    };
    
    // Get only visible columns in their current order
    const visibleColumns = currentColumnOrder.filter(col => columnVisibility[col]);
    
    // Map data according to visible columns and their order
    const data = communesData.map(commune => {
        const row = {};
        visibleColumns.forEach(col => {
            row[columnDisplayNames[col] || col] = commune[col];
        });
        return row;
    });

    exportToCSV(data);
}
