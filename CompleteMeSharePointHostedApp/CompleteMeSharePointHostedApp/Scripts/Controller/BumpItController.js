myApp.controller('DropdownCtrl', ['$scope', function ($scope) {
    var d = moment();
    $scope.items = [
       { text: "Tomorrow", date: moment(d).add('d', 1).format() },
       { text: "Tomorrow +1", date: moment(d).add('d', 2).format() },
       { text: "Tomorrow +3", date: moment(d).add('d', 3).format() },
       { text: "Next Monday", date: moment(d).day(8).format() },// next Monday,
       { text: "Next Friday", date: moment(d).day(12).format() },// next Friday,
       { text: "Two Weeks", date: moment(d).add('w', 2).format() },// two weeks
       { text: "Three Weeks", date: moment(d).add('w', 3).format() },// three weeks
       { text: "Four Weeks", date: moment(d).add('w', 4).format() },// four weeks
       { text: "Forget about it", date: moment(d).add('M', 6).format() },
       { text: "I'll make you an offer you can't refuse", date: moment(d).add('y', 1).format() }
    ];
}]);