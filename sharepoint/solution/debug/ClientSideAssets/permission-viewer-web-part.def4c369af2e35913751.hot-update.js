"use strict";
self["webpackHotUpdate_76d5422e_a579_4a00_99b3_192667104868_0_0_1"]("permission-viewer-web-part",{

/***/ 89135:
/*!**********************************************************************!*\
  !*** ./lib/webparts/permissionViewer/components/PermissionViewer.js ***!
  \**********************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 85959);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! @fluentui/react/lib/Pivot */ 92070);
/* harmony import */ var _fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! @fluentui/react/lib/Pivot */ 15369);
/* harmony import */ var _fluentui_react_lib_SearchBox__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! @fluentui/react/lib/SearchBox */ 21262);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 5613);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 29425);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 80954);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 49885);
Object(function webpackMissingModule() { var e = new Error("Cannot find module '../services/PermissionService'"); e.code = 'MODULE_NOT_FOUND'; throw e; }());
/* harmony import */ var _Header__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Header */ 26783);
/* harmony import */ var _StatsCards__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./StatsCards */ 54088);
/* harmony import */ var _SitePermissions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./SitePermissions */ 82561);
/* harmony import */ var _ListPermissions__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./ListPermissions */ 20548);
/* harmony import */ var _LoadingState__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./LoadingState */ 37087);
/* harmony import */ var _DeepScanDialog__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./DeepScanDialog */ 34937);
/* harmony import */ var _utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../utils/CsvExport */ 2759);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! @fluentui/react */ 4312);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! @fluentui/react */ 10548);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! @fluentui/react */ 87295);
/* harmony import */ var _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./PermissionViewer.module.scss */ 29929);
















var PermissionViewer = function (props) {
    var _a;
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState('site'), activeTab = _b[0], setActiveTab = _b[1];
    var _c = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), sitePermissions = _c[0], setSitePermissions = _c[1];
    var _d = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), filteredSitePermissions = _d[0], setFilteredSitePermissions = _d[1];
    var _e = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), lists = _e[0], setLists = _e[1];
    var _f = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), filteredLists = _f[0], setFilteredLists = _f[1];
    var _g = react__WEBPACK_IMPORTED_MODULE_0__.useState({ totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 }), stats = _g[0], setStats = _g[1];
    var _h = react__WEBPACK_IMPORTED_MODULE_0__.useState(true), isLoading = _h[0], setIsLoading = _h[1];
    var _j = react__WEBPACK_IMPORTED_MODULE_0__.useState('Loading site permissions...'), loadingMessage = _j[0], setLoadingMessage = _j[1];
    var _k = react__WEBPACK_IMPORTED_MODULE_0__.useState(), permissionService = _k[0], setPermissionService = _k[1];
    var _l = react__WEBPACK_IMPORTED_MODULE_0__.useState(''), searchText = _l[0], setSearchText = _l[1];
    var _m = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isExporting = _m[0], setIsExporting = _m[1];
    var _o = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isScanning = _o[0], setIsScanning = _o[1];
    // Deep Scan State
    var _p = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isDeepScanOpen = _p[0], setIsDeepScanOpen = _p[1];
    var _q = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), deepScanItems = _q[0], setDeepScanItems = _q[1];
    var _r = react__WEBPACK_IMPORTED_MODULE_0__.useState(''), deepScanListTitle = _r[0], setDeepScanListTitle = _r[1];
    var _s = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), confirmScanList = _s[0], setConfirmScanList = _s[1];
    var _t = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), scanNoResults = _t[0], setScanNoResults = _t[1];
    var _u = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), errorMessage = _u[0], setErrorMessage = _u[1];
    // Delete Confirmation State
    var _v = react__WEBPACK_IMPORTED_MODULE_0__.useState({ isOpen: false, title: '', subText: '', onConfirm: function () { } }), deleteConfirmState = _v[0], setDeleteConfirmState = _v[1];
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        var service = new Object(function webpackMissingModule() { var e = new Error("Cannot find module '../services/PermissionService'"); e.code = 'MODULE_NOT_FOUND'; throw e; }())(props.spHttpClient, props.webUrl);
        setPermissionService(service);
        loadData(service);
    }, [props.excludedLists]);
    var loadData = function (service) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var sitePerms, siteStats, listsData, uniqueListsCount;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setIsLoading(true);
                    setLoadingMessage('Loading site permissions...');
                    return [4 /*yield*/, service.getSiteRoleAssignments()];
                case 1:
                    sitePerms = _a.sent();
                    setSitePermissions(sitePerms);
                    setFilteredSitePermissions(sitePerms);
                    return [4 /*yield*/, service.getSiteStats()];
                case 2:
                    siteStats = _a.sent();
                    // Load lists
                    setLoadingMessage('Loading lists and libraries...');
                    return [4 /*yield*/, service.getLists(props.excludedLists)];
                case 3:
                    listsData = _a.sent();
                    uniqueListsCount = listsData.filter(function (l) { return l.HasUniqueRoleAssignments; }).length;
                    setStats((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, siteStats), { uniquePermissionsCount: uniqueListsCount + (sitePerms.length > 0 ? 1 : 0) // Simplified logic: +1 if site has perms (it always does)
                     }));
                    // Sort lists: Unique first, then Inherited
                    listsData.sort(function (a, b) {
                        if (a.HasUniqueRoleAssignments === b.HasUniqueRoleAssignments)
                            return 0;
                        return a.HasUniqueRoleAssignments ? -1 : 1;
                    });
                    setLists(listsData);
                    setFilteredLists(listsData);
                    setIsLoading(false);
                    return [2 /*return*/];
            }
        });
    }); };
    var handleRefresh = function () {
        if (permissionService) {
            loadData(permissionService);
        }
    };
    var onSearch = function (newValue) {
        setSearchText(newValue);
        var lower = newValue.toLowerCase();
        if (activeTab === 'site') {
            var filtered = sitePermissions.filter(function (p) {
                var _a;
                return p.Member.Title.toLowerCase().includes(lower) ||
                    ((_a = p.Member.Email) === null || _a === void 0 ? void 0 : _a.toLowerCase().includes(lower));
            });
            setFilteredSitePermissions(filtered);
        }
        else {
            var filtered = lists.filter(function (l) {
                return l.Title.toLowerCase().includes(lower) ||
                    l.ServerRelativeUrl.toLowerCase().includes(lower);
            });
            setFilteredLists(filtered);
        }
    };
    var handleGetListPermissions = function (listId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var list;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/, []];
                    list = lists.find(function (l) { return l.Id === listId; });
                    return [4 /*yield*/, permissionService.getListRoleAssignments(listId, list ? list.Title : '')];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    }); };
    var handleRemoveSitePermission = function (principalId, principalName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove Permissions?",
            subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', "? This will remove all permissions for this user on this site."),
            onConfirm: function () { return executeRemoveSitePermission(principalId); },
            onCancel: function () { }
        });
    };
    var executeRemoveSitePermission = function (principalId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var success, sitePerms, lower_1, filtered, error_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); });
                    setIsLoading(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    return [4 /*yield*/, permissionService.removeSitePermission(principalId)];
                case 2:
                    success = _a.sent();
                    if (!success) return [3 /*break*/, 4];
                    return [4 /*yield*/, permissionService.getSiteRoleAssignments()];
                case 3:
                    sitePerms = _a.sent();
                    setSitePermissions(sitePerms);
                    if (searchText) {
                        lower_1 = searchText.toLowerCase();
                        filtered = sitePerms.filter(function (p) {
                            var _a;
                            return p.Member.Title.toLowerCase().includes(lower_1) ||
                                ((_a = p.Member.Email) === null || _a === void 0 ? void 0 : _a.toLowerCase().includes(lower_1));
                        });
                        setFilteredSitePermissions(filtered);
                    }
                    else {
                        setFilteredSitePermissions(sitePerms);
                    }
                    return [3 /*break*/, 5];
                case 4:
                    setErrorMessage("Failed to remove permission. Please try again or check console for details.");
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_1 = _a.sent();
                    console.error("Error removing site permission", error_1);
                    setErrorMessage("An unexpected error occurred while removing permission.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsLoading(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    var handleRemoveListPermission = function (listId, principalId, principalName) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    setDeleteConfirmState({
                        isOpen: true,
                        title: "Remove Permissions?",
                        subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', " on this list?"),
                        onConfirm: function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
                            var success, error_2;
                            return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); });
                                        if (!permissionService) {
                                            resolve(false);
                                            return [2 /*return*/];
                                        }
                                        _a.label = 1;
                                    case 1:
                                        _a.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, permissionService.removeListPermission(listId, principalId)];
                                    case 2:
                                        success = _a.sent();
                                        if (!success) {
                                            setErrorMessage("Failed to remove list permission. Please try again or check console for details.");
                                        }
                                        resolve(success);
                                        return [3 /*break*/, 4];
                                    case 3:
                                        error_2 = _a.sent();
                                        console.error("Error removing list permission", error_2);
                                        setErrorMessage("An unexpected error occurred while removing list permission.");
                                        resolve(false);
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); },
                        onCancel: function () {
                            resolve(false);
                        }
                    });
                })];
        });
    }); };
    var handleRemoveDeepScanItemPermission = function (itemId, principalId, principalName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove Permissions?",
            subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', " on this item?"),
            onConfirm: function () { return executeRemoveDeepScanItemPermission(itemId, principalId); },
            onCancel: function () { }
        });
    };
    var executeRemoveDeepScanItemPermission = function (itemId, principalId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var list, success, error_3;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService || !deepScanListTitle)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); });
                    list = lists.find(function (l) { return l.Title === deepScanListTitle; });
                    if (!list)
                        return [2 /*return*/];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, permissionService.removeItemPermission(list.Id, itemId, principalId)];
                case 2:
                    success = _a.sent();
                    if (success) {
                        setDeepScanItems(function (prevItems) {
                            return prevItems.map(function (item) {
                                if (item.Id === itemId) {
                                    var newRoles = item.RoleAssignments.filter(function (ra) { return ra.PrincipalId !== principalId; });
                                    if (newRoles.length === 0) {
                                        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, item), { RoleAssignments: [] });
                                    }
                                    return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, item), { RoleAssignments: newRoles });
                                }
                                return item;
                            }).filter(function (item) { return item !== null; });
                        });
                    }
                    else {
                        setErrorMessage("Failed to remove item permission.");
                    }
                    return [3 /*break*/, 4];
                case 3:
                    error_3 = _a.sent();
                    console.error("Error removing item permission", error_3);
                    setErrorMessage("An unexpected error occurred while removing item permission.");
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var handleExport = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var error_4;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setIsExporting(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    if (!(activeTab === 'site')) return [3 /*break*/, 3];
                    return [4 /*yield*/, (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__.exportSitePermissions)(filteredSitePermissions, permissionService)];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__.exportListPermissions)(filteredLists, permissionService)];
                case 4:
                    _a.sent();
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_4 = _a.sent();
                    console.error("Export failed", error_4);
                    alert("Export failed. See console for details.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsExporting(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    // ...
    var handleDeepScan = function (listId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var list;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            list = lists.find(function (l) { return l.Id === listId; });
            if (!list)
                return [2 /*return*/];
            setScanNoResults(false);
            setConfirmScanList({ id: list.Id, title: list.Title });
            return [2 /*return*/];
        });
    }); };
    var executeDeepScan = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var listId, listTitle, items, e_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService || !confirmScanList)
                        return [2 /*return*/];
                    listId = confirmScanList.id;
                    listTitle = confirmScanList.title;
                    setIsScanning(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, permissionService.getUniquePermissionItems(listId)];
                case 2:
                    items = _a.sent();
                    if (items.length > 0) {
                        setDeepScanItems(items);
                        setDeepScanListTitle(listTitle);
                        setIsDeepScanOpen(true);
                        setConfirmScanList(null); // Close confirm dialog only if opening results
                    }
                    else {
                        setScanNoResults(true);
                        // Keep dialog open to show message
                    }
                    return [3 /*break*/, 5];
                case 3:
                    e_1 = _a.sent();
                    console.error(e_1);
                    alert("Error during deep scan.");
                    setConfirmScanList(null); // Close on error too
                    return [3 /*break*/, 5];
                case 4:
                    setIsScanning(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var downloadDeepScanResults = function () {
        (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__.exportDeepScanResults)(deepScanItems, deepScanListTitle);
    };
    var getDialogTitle = function () {
        if (scanNoResults)
            return 'Deep Scan Complete';
        if (isScanning)
            return 'Deep Scan in Progress...';
        return 'Start Deep Scan?';
    };
    var getDialogSubText = function () {
        if (scanNoResults) {
            return "No items with unique permissions were found in \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\". All items inherit permissions.");
        }
        if (isScanning) {
            return "Scanning \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\". This may take a few moments depending on the number of items...");
        }
        return "This will verify every single item in \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\" to find unique permissions. This might take a while for large lists. Continue?");
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].permissionViewer },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].webpartContainer },
            (props.showComponentHeader !== false) && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_Header__WEBPACK_IMPORTED_MODULE_2__.Header, { onRefresh: handleRefresh, isLoading: isLoading, themeVariant: props.themeVariant, opacity: (_a = props.headerOpacity) !== null && _a !== void 0 ? _a : 100, title: props.webPartTitle, titleFontSize: props.webPartTitleFontSize })),
            (props.showStats !== false) && react__WEBPACK_IMPORTED_MODULE_0__.createElement(_StatsCards__WEBPACK_IMPORTED_MODULE_3__.StatsCards, { stats: stats }),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].tabsContainer },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_11__.Pivot, { onLinkClick: function (item) {
                        if (item === null || item === void 0 ? void 0 : item.props.itemKey) {
                            setActiveTab(item.props.itemKey);
                        }
                    }, selectedKey: activeTab },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_12__.PivotItem, { headerText: "Site Permissions", itemKey: "site", itemIcon: "Shield" }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_12__.PivotItem, { headerText: "Lists & Libraries", itemKey: "lists", itemIcon: "List" }))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].toolbar },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_SearchBox__WEBPACK_IMPORTED_MODULE_13__.SearchBox, { placeholder: "Search users or groups...", onChange: function (_, newValue) { return onSearch(newValue || ''); }, value: searchText, className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].searchBox }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__.DefaultButton, { text: isExporting ? "Exporting..." : "Export to CSV", iconProps: { iconName: 'Download' }, onClick: handleExport, disabled: isExporting || isLoading, className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].exportBtn, styles: {
                        root: { height: '32px' }, // Maintain height
                        label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                    } })),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].content },
                isLoading && react__WEBPACK_IMPORTED_MODULE_0__.createElement(_LoadingState__WEBPACK_IMPORTED_MODULE_6__.LoadingState, { message: loadingMessage }),
                !isLoading && activeTab === 'site' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SitePermissions__WEBPACK_IMPORTED_MODULE_4__.SitePermissions, { permissions: filteredSitePermissions, permissionService: permissionService, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveSitePermission })),
                !isLoading && activeTab === 'lists' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_ListPermissions__WEBPACK_IMPORTED_MODULE_5__.ListPermissions, { lists: filteredLists, getListPermissions: handleGetListPermissions, onScanItems: handleDeepScan, themeVariant: props.themeVariant, buttonFontSize: props.buttonFontSize, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveListPermission })))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_DeepScanDialog__WEBPACK_IMPORTED_MODULE_7__.DeepScanDialog, { isOpen: isDeepScanOpen, onDismiss: function () { return setIsDeepScanOpen(false); }, listTitle: deepScanListTitle, items: deepScanItems, onDownload: downloadDeepScanResults, buttonFontSize: props.buttonFontSize, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveDeepScanItemPermission }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.Dialog, { hidden: !confirmScanList, onDismiss: function () { if (!isScanning)
                setConfirmScanList(null); }, dialogContentProps: {
                type: _fluentui_react__WEBPACK_IMPORTED_MODULE_16__.DialogType.normal,
                title: getDialogTitle(),
                subText: getDialogSubText()
            } }, isScanning ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_17__.Spinner, { size: _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_18__.SpinnerSize.large, label: "Scanning for unique permissions..." })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.DialogFooter, null, scanNoResults ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: function () { return setConfirmScanList(null); }, text: "Close", styles: {
                root: { height: '32px' },
                label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
            } })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(react__WEBPACK_IMPORTED_MODULE_0__.Fragment, null,
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: executeDeepScan, text: "Start Scan", styles: {
                    root: { height: '32px' },
                    label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                } }),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__.DefaultButton, { onClick: function () { return setConfirmScanList(null); }, text: "Cancel", styles: {
                    root: { height: '32px' },
                    label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                } })))))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.Dialog, { hidden: !deleteConfirmState.isOpen, onDismiss: function () { return setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); }); }, dialogContentProps: {
                type: _fluentui_react__WEBPACK_IMPORTED_MODULE_16__.DialogType.normal,
                title: deleteConfirmState.title,
                subText: deleteConfirmState.subText,
            }, modalProps: {
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
            } },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.DialogFooter, null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: deleteConfirmState.onConfirm, text: "Remove", styles: {
                        root: { background: '#d13438', border: '1px solid #d13438' }, // Red color for danger
                        rootHovered: { background: '#a4262c' },
                        label: { fontWeight: 600 }
                    } }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__.DefaultButton, { onClick: function () { return setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); }); }, text: "Cancel" }))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.Dialog, { hidden: !errorMessage, onDismiss: function () { return setErrorMessage(null); }, dialogContentProps: {
                type: _fluentui_react__WEBPACK_IMPORTED_MODULE_16__.DialogType.normal,
                title: 'Error',
                subText: errorMessage || 'An unexpected error occurred.'
            } },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.DialogFooter, null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: function () { return setErrorMessage(null); }, text: "OK" })))));
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (PermissionViewer);


/***/ }),

/***/ 82561:
/*!*********************************************************************!*\
  !*** ./lib/webparts/permissionViewer/components/SitePermissions.js ***!
  \*********************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   SitePermissions: () => (/* binding */ SitePermissions)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 85959);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _UserPersona__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./UserPersona */ 31157);
/* harmony import */ var _PermissionBadge__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./PermissionBadge */ 21146);
/* harmony import */ var _fluentui_react_lib_DetailsList__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! @fluentui/react/lib/DetailsList */ 79370);
/* harmony import */ var _fluentui_react_lib_DetailsList__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! @fluentui/react/lib/DetailsList */ 74423);
/* harmony import */ var _fluentui_react_lib_DetailsList__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! @fluentui/react/lib/DetailsList */ 37805);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 44533);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 80954);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 49885);
/* harmony import */ var _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./PermissionViewer.module.scss */ 29929);








var UserGroupCell = function (_a) {
    var item = _a.item, expandedGroups = _a.expandedGroups, onToggle = _a.onToggle, fontSize = _a.fontSize, onRemovePermission = _a.onRemovePermission;
    var isGroup = item.Member.PrincipalType === 8 || item.Member.PrincipalType === 4;
    var isUser = item.Member.PrincipalType === 1;
    var isExpanded = expandedGroups.has(item.Member.Id);
    var depth = item.depth || 0;
    var handleDelete = function () {
        if (onRemovePermission) {
            onRemovePermission(item.Member.Id, item.Member.Title);
        }
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { style: { paddingLeft: "".concat(depth * 24, "px"), display: 'flex', alignItems: 'center', fontSize: fontSize, justifyContent: 'space-between' } },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { style: { display: 'flex', alignItems: 'center' } },
            isGroup && depth === 0 && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_4__.IconButton, { iconProps: { iconName: isExpanded ? 'ChevronDown' : 'ChevronRight' }, onClick: function () { return onToggle(item); }, styles: { root: { height: 24, width: 24, marginRight: 4 } } })),
            item.isLoading ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: '8px', fontStyle: 'italic', color: '#605e5c' } },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_5__.Spinner, { size: _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_6__.SpinnerSize.xSmall }),
                "Loading...")) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_UserPersona__WEBPACK_IMPORTED_MODULE_1__.UserPersona, { user: item.Member, secondaryText: depth > 0 ? 'Member' : undefined, fontSize: fontSize }))),
        isUser && depth === 0 && onRemovePermission && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_4__.IconButton, { iconProps: { iconName: 'Delete' }, title: "Remove Permission", onClick: handleDelete, styles: { root: { height: 24, width: 24, color: '#a80000' } } }))));
};
var PrincipalTypeCell = function (_a) {
    var item = _a.item, fontSize = _a.fontSize;
    var typeMap = { 1: 'User', 4: 'Security Group', 8: 'SharePoint Group' };
    return react__WEBPACK_IMPORTED_MODULE_0__.createElement("span", { style: { fontSize: fontSize } }, typeMap[item.Member.PrincipalType] || 'Unknown');
};
var PermissionLevelCell = function (_a) {
    var item = _a.item, fontSize = _a.fontSize;
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { style: { display: 'flex', gap: '8px', flexWrap: 'wrap' } }, item.RoleDefinitionBindings.map(function (role) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_PermissionBadge__WEBPACK_IMPORTED_MODULE_2__.PermissionBadge, { key: role.Id, permission: role.Name, fontSize: fontSize ? "calc(".concat(fontSize, " - 2px)") : undefined })); })));
};
var renderPrincipalType = function (item, fontSize) { return react__WEBPACK_IMPORTED_MODULE_0__.createElement(PrincipalTypeCell, { item: item, fontSize: fontSize }); };
var renderPermissionLevel = function (item, fontSize) { return react__WEBPACK_IMPORTED_MODULE_0__.createElement(PermissionLevelCell, { item: item, fontSize: fontSize }); };
var SitePermissions = function (props) {
    var permissions = props.permissions, permissionService = props.permissionService, onRemovePermission = props.onRemovePermission;
    var _a = react__WEBPACK_IMPORTED_MODULE_0__.useState(new Set()), expandedGroups = _a[0], setExpandedGroups = _a[1];
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState({}), groupMembers = _b[0], setGroupMembers = _b[1];
    var _c = react__WEBPACK_IMPORTED_MODULE_0__.useState(new Set()), loadingGroups = _c[0], setLoadingGroups = _c[1];
    var toggleGroup = function (group) { return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__awaiter)(void 0, void 0, void 0, function () {
        var groupId, newExpanded, users, memberAssignments_1, err_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_7__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    groupId = group.Member.Id;
                    newExpanded = new Set(expandedGroups);
                    if (newExpanded.has(groupId)) {
                        newExpanded.delete(groupId);
                        setExpandedGroups(newExpanded);
                        return [2 /*return*/];
                    }
                    newExpanded.add(groupId);
                    setExpandedGroups(newExpanded);
                    if (!(!groupMembers[groupId] && permissionService)) return [3 /*break*/, 5];
                    setLoadingGroups(function (prev) { return new Set(prev).add(groupId); });
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, permissionService.getGroupMembers(groupId)];
                case 2:
                    users = _a.sent();
                    memberAssignments_1 = users.map(function (u) { return ({
                        Member: u,
                        PrincipalId: u.Id,
                        RoleDefinitionBindings: group.RoleDefinitionBindings // Inherit permission levels from group
                    }); });
                    setGroupMembers(function (prev) {
                        var _a;
                        return ((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)({}, prev), (_a = {}, _a[groupId] = memberAssignments_1, _a)));
                    });
                    return [3 /*break*/, 5];
                case 3:
                    err_1 = _a.sent();
                    console.error(err_1);
                    return [3 /*break*/, 5];
                case 4:
                    setLoadingGroups(function (prev) {
                        var next = new Set(prev);
                        next.delete(groupId);
                        return next;
                    });
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var displayItems = react__WEBPACK_IMPORTED_MODULE_0__.useMemo(function () {
        var items = [];
        permissions.forEach(function (p) {
            items.push((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)({}, p), { depth: 0 }));
            var groupId = p.Member.Id;
            if ((p.Member.PrincipalType === 8 || p.Member.PrincipalType === 4) && expandedGroups.has(groupId)) {
                if (loadingGroups.has(groupId)) {
                    // Placeholder for loading
                    items.push({
                        Member: { Id: -1, Title: 'Loading...', IsHiddenInUI: false, PrincipalType: 1 },
                        PrincipalId: -1,
                        RoleDefinitionBindings: [],
                        depth: 1,
                        isLoading: true
                    });
                }
                else if (groupMembers[groupId]) {
                    groupMembers[groupId].forEach(function (m) {
                        items.push((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_7__.__assign)({}, m), { depth: 1 }));
                    });
                }
            }
        });
        return items;
    }, [permissions, expandedGroups, groupMembers, loadingGroups]);
    var renderUserGroup = react__WEBPACK_IMPORTED_MODULE_0__.useCallback(function (item) { return (react__WEBPACK_IMPORTED_MODULE_0__.createElement(UserGroupCell, { item: item, expandedGroups: expandedGroups, onToggle: toggleGroup, fontSize: props.contentFontSize, onRemovePermission: onRemovePermission })); }, [expandedGroups, toggleGroup, props.contentFontSize, onRemovePermission]);
    var columns = react__WEBPACK_IMPORTED_MODULE_0__.useMemo(function () { return [
        {
            key: 'user',
            name: 'User/Group',
            fieldName: 'Member',
            minWidth: 200,
            maxWidth: 400,
            onRender: renderUserGroup
        },
        {
            key: 'type',
            name: 'Type',
            fieldName: 'Member',
            minWidth: 100,
            maxWidth: 150,
            onRender: function (item) { return renderPrincipalType(item, props.contentFontSize); }
        },
        {
            key: 'level',
            name: 'Permission Level',
            fieldName: 'RoleDefinitionBindings',
            minWidth: 150,
            maxWidth: 200,
            onRender: function (item) { return renderPermissionLevel(item, props.contentFontSize); }
        }
    ]; }, [renderUserGroup, props.contentFontSize]);
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].content },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_3__["default"].permissionTable, style: { border: '1px solid #e1dfdd', borderRadius: '4px', overflow: 'hidden' } },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_DetailsList__WEBPACK_IMPORTED_MODULE_8__.DetailsList, { items: displayItems, columns: columns, selectionMode: _fluentui_react_lib_DetailsList__WEBPACK_IMPORTED_MODULE_9__.SelectionMode.none, layoutMode: _fluentui_react_lib_DetailsList__WEBPACK_IMPORTED_MODULE_10__.DetailsListLayoutMode.justified, styles: {
                    root: { background: '#ffffff', fontSize: props.contentFontSize },
                    headerWrapper: { background: '#faf9f8' },
                    contentWrapper: { fontSize: props.contentFontSize }
                } }))));
};


/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ (() => {
/******/ 	__webpack_require__.h = () => ("ca2d52bd176eecd06765")
/******/ })();
/******/ 
/******/ }
);
//# sourceMappingURL=permission-viewer-web-part.def4c369af2e35913751.hot-update.js.map