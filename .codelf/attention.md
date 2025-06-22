## Development Guidelines

### Framework and Language
> Analyze the framework and language choices for this project, focusing on best practices and standardization.

**Framework Considerations:**
- Version Compatibility: Ensure all dependencies are compatible with the chosen framework version
- Feature Usage: Leverage framework-specific features rather than reinventing solutions
- Performance Patterns: Follow recommended patterns for optimal performance
- Upgrade Strategy: Plan for future framework updates with minimal disruption
- Importance Notes for Framework: 
	* {important notes}

**Language Best Practices:**
- Type Safety: Use strong typing where available to prevent runtime errors
- Modern Features: Utilize modern language features while maintaining compatibility
- Consistency: Apply consistent coding patterns throughout the codebase
- Documentation: Document language-specific implementations and workarounds

### Code Abstraction and Reusability
> During development, prioritize code abstraction and reusability to ensure modular and component-based functionality. Try to search for existing solutions before reinventing the wheel.
> List below the directory structure of common components, utility functions, and API encapsulations in the current project.


**Modular Design Principles:**
- Single Responsibility: Each module is responsible for only one functionality
- High Cohesion, Low Coupling: Related functions are centralized, reducing dependencies between modules
- Stable Interfaces: Expose stable interfaces externally while internal implementations can vary

**Reusable Component Library:**
e.g.
```
root
- pkg 
    - utils
		- time // include time related functions: getNow, getToday, getYesterday, etc.
    - views
```

### Coding Standards and Tools
**Code Formatting Tools:**
- [ESLint (version)]() // JavaScript/TypeScript code checking
- [Prettier (version)]() // Code formatting
- [StyleLint (version)]() // CSS/SCSS code checking

**Naming and Structure Conventions:**
- Semantic Naming: Variable/function names should clearly express their purpose
- Consistent Naming Style: Frontend uses camelCase, CSS uses kebab-case
- Directory Structure follows functional responsibility division

### Frontend-Backend Collaboration Standards
**API Design and Documentation:**
- RESTful design principles
	* Use HTTP methods (GET, POST, PUT, DELETE) to represent operations
	...
- Timely interface documentation updates
	* Document API endpoints, parameters, and responses
	...
- Unified error handling specifications
	...
	

**Data Flow:**
- Clear frontend state management
	* Use a state management library (e.g., Redux, Pinia) for consistent state handling
	...
- Data validation on both frontend and backend
	* Validate data types and constraints
	...
- Standardized asynchronous operation handling
	* Use consistent API call patterns
	...

### Performance and Security
**Performance Optimization Focus:**
- Resource loading optimization
	* Use code splitting and lazy loading
	...
- Rendering performance optimization
	* Use virtualization and pagination for large lists
	...
- Appropriate use of caching
	* Implement caching strategies for API responses
	...

**Security Measures:**
- Input validation and filtering
	* Validate user inputs and sanitize data
	...
- Protection of sensitive information
	* Use secure authentication and authorization mechanisms
	...
- Access control mechanisms
	* Implement role-based access control
	...