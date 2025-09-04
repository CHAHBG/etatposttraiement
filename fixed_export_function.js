/**
 * Exports data to specified format
 * @param {'csv'} format
 */
function exportData(format = 'csv') {
    const data = communesData.map(commune => ({
        Commune: commune.commune,
        Région: commune.region,
        'Total Parcelles': commune.totalParcelles,
        '% du Total': commune.percentTotal,
        NICAD: commune.nicad,
        '% NICAD': commune.percentNicad,
        CTASF: commune.ctasf,
        '% CTASF': commune.percentCtasf,
        Délibérées: commune.deliberated,
        '% Délibérée': commune.percentDeliberated,
        'Parcelles brutes': commune.parcellesBrutes,
        'Parcelles collectées (sans doublon géométrique)': commune.collected,
        'Parcelles enquêtées': commune.surveyed,
        'Motifs de rejet post-traitement': commune.rejectionReasons,
        'Parcelles retenues après post-traitement': commune.retained,
        'Parcelles validées par l\'URM': commune.validated,
        'Parcelles rejetées par l\'URM': commune.rejected,
        'Motifs de rejet URM': commune.urmRejectionReasons,
        'Parcelles corrigées': commune.corrected,
        Geomaticien: commune.geomaticien,
        'Parcelles individuelles jointes': commune.individualJoined,
        'Parcelles collectives jointes': commune.collectiveJoined,
        'Parcelles non jointes': commune.unjoined,
        'Doublons supprimés': commune.duplicatesRemoved,
        'Taux suppression doublons (%)': commune.duplicateRemovalRate,
        'Parcelles en conflit': commune.parcelsInConflict,
        'Significant Duplicates': commune.significantDuplicates,
        'Parc. post-traitées lot 1-46': commune.postProcessedLot1_46,
        'Statut jointure': commune.jointureStatus,
        'Message d\'erreur jointure': commune.jointureErrorMessage
    }));

    if (format === 'csv') {
        exportToCSV(data);
    }
}
